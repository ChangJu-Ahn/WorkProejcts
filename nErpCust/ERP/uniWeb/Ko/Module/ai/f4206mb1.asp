<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : �ڱ� 
'*  3. Program ID           : f4204mb1
'*  4. Program Name         : �������Աݵ�� 
'*  5. Program Desc         :
'*  6. Comproxy List        : FL0061, FL0069
'*  7. Modified date(First) : 2002/05/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Oh, Soo Min
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :																	
'**********************************************************************************************



	On Error Resume Next                                                            '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status    
    
    Call HideStatusWnd                                                              '��: Hide Processing message
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
	Dim lgOpModeCRUD
	gChangeOrgId = GetGlobalInf("gChangeOrgId")
    '---------------------------------------Common-----------------------------------------------------------
                                                          '��: Set to space
	lgOpModeCRUD	= Request("txtMode")					'�� : ���� ���¸� ���� 

	'------ Developer Coding part (Start ) ------------------------------------------------------------------Dim strMode

	Dim strLoanNo

	Dim iStrPrevKey
	Dim iIntMaxRows
	Dim iIntQueryCount

    iStrPrevKey		= Trim(Request("lgStrPrevKey"))        
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    
	'------ Developer Coding part (End   ) ------------------------------------------------------------------     
Select Case lgOpModeCRUD
     Case CStr(UID_M0001)                                                         '��: Query
          Call SubBizQuery()

     Case CStr(UID_M0002)       
          Call SubBizSave()

     Case CStr(UID_M0003)                                                         '��: Delete
          Call SubBizDelete()
     Case "PAFG400"		'��ȯ���� 
		  Call SubPlanExec()
End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                                 '��: Protect system from crashing
    Err.Clear                                                                            '��: Clear Error status
    
    Dim strLoanTermYr
    Dim strLoanTermMnth
    
	Const L_loan_no = 0
    Const L_loan_basic_fg = 1
    Const L_loan_fg = 2
    Const L_loan_plc_type = 3   

	Const C_LOAN_NO = 0
    Const C_LOAN_NM = 1
    Const C_LOAN_FG = 2
    Const C_LOAN_DT = 3
    Const C_DOC_CUR = 4
    Const C_XCH_RATE = 5
    Const C_LOAN_AMT = 6
    Const C_LOAN_LOC_AMT = 7
    Const C_RDP_AMT = 8
    Const C_RDP_LOC_AMT = 9
    Const C_INT_PAY_AMT = 10
    Const C_INT_PAY_LOC_AMT = 11
    Const C_BAS_RDP_AMT = 12
    Const C_BAS_RDP_LOC_AMT = 13
    Const C_BAS_INT_PAY_AMT = 14
    Const C_BAS_INT_PAY_LOC_AMT = 15
    Const C_LOAN_BAL_AMT = 16
    Const C_LOAN_BAL_LOC_AMT = 17
    Const C_LOAN_TYPE = 18
    Const C_LOAN_TYPE_NM = 19
    Const C_DUE_DT = 20
    Const C_RCPT_TYPE = 21
    Const C_RCPT_NM = 22
    Const C_LOAN_INT_RATE = 23
    Const C_LOAN_TERM = 24
    Const C_PR_RDP_TIMES = 25
    Const C_PR_RDP_UNIT_AMT = 26
    Const C_PR_RDP_UNIT_LOC_AMT = 27
    Const C_ST_PR_RDP_DT = 28
    Const C_ST_INT_DUE_DT = 29
    Const C_PR_RDP_PERD = 30
    Const C_PR_RDP_COND = 31
    Const C_INT_PAY_PERD = 32
    Const C_INT_PAY_PERD_BASE = 33
    Const C_INT_PAY_STND = 34
    Const C_INT_BASE_MTHD = 35
    Const C_INT_VOTL = 36
    Const C_DAY_MTHD = 37
    Const C_RDP_CLS_FG = 38
    Const C_LC_DOC_NO = 39
    Const C_GL_NO = 40
    Const C_TEMP_GL_NO = 41
    Const C_PR_RDP_DT = 42
    Const C_INT_PAY_DT = 43
    Const C_LOAN_BASIC_FG = 44
    Const C_LOAN_PLC_TYPE = 45
    Const C_LOAN_BASIC_DT = 46
    Const C_LOAN_DESC = 47
    Const C_ST_ADV_INT_PAY_AMT = 48
    Const C_ST_ADV_INT_PAY_LOC_AMT = 49
    Const C_CLS_RO_FG = 50
    Const C_CONF_FG = 51
    Const C_BP_CD = 52
    Const C_RDP_SPRD_FG = 53
    Const C_LOAN_BANK_CD = 54
    Const C_LOAN_BANK_NM = 55
    Const C_BANK_CD = 56
    Const C_BANK_NM = 57
    Const C_BANK_ACCT_NO = 58
    Const C_ORG_CHANGE_ID = 59
    Const C_DEPT_CD = 60
    Const C_DEPT_NM = 61
    Const C_INTERNAL_CD = 62
    Const C_BP_NM = 63
    Const C_LOAN_ACCT_CD = 64		'���Աݰ����ڵ� 
    Const C_LOAN_ACCT_NM = 65		'���Աݰ����ڵ� 
    Const C_RCPT_ACCT_CD = 66		'�Ա����������ڵ� 
    Const C_RCPT_ACCT_NM = 67		'�Ա����������ڵ� 
    Const C_INT_ACCT_CD = 68		'�������ް����ڵ� 
    Const C_INT_ACCT_NM = 69		'�������ް����ڵ� 
	Const C_USER_FIELD1 = 70
	Const C_USER_FIELD2 = 71

	Dim iPAFG405list																'�� : ��ȸ�� Component Dll ��� ���� 

	Dim iCommandSent
	Dim I_f_ln_info	
	Dim E_f_ln_info
	Dim E_PrevNext_Code
	
	ReDim I_f_ln_info(L_loan_plc_type)

    iCommandSent = Request("txtCommand")
    I_f_ln_info(L_loan_no)		 = Trim(Request("txtLoanNo"))
    I_f_ln_info(L_loan_basic_fg) = Trim(Request("txtLoanBasicFg"))
    I_f_ln_info(L_loan_fg)		 = ""
    I_f_ln_info(L_loan_plc_type) = Trim(Request("txtLoanPlcType"))            

    Set iPAFG405list = server.CreateObject("PAFG405.bFLkUpLnSvr")

    If CheckSYSTEMError(Err,True) = True Then		
		Exit Sub
    End If    
    
    Call iPAFG405list.F_LOOKUP_LN_SVR(gStrGlobalCollection, iCommandSent, I_f_ln_info, E_f_ln_info, E_PrevNext_Code )										
										
    If CheckSYSTEMError(Err,True) = True Then
		Set iPAFG405list = nothing
		Exit Sub
    End If
	Set iPAFG405list = nothing 

	If Trim(E_PrevNext_Code(0)) = "900011" Or Trim(E_PrevNext_Code(0)) = "900012" Then
		Call DisplayMsgBox(E_PrevNext_Code(0), VbOKOnly, "", "", I_MKSCRIPT)
	End If
    
    Response.Write "<Script Language=vbscript>  " & vbCr
    Response.Write " with parent.frm1"			  & vbCr     
    Response.Write " .txtLoanNo.value       = """ & ConvSPChars(E_f_ln_info(C_LOAN_NO)) & """              " & vbCr
    Response.Write " .txtLoanNm.value       = """ & ConvSPChars(E_f_ln_info(C_LOAN_NM)) & """              " & vbCr                                            '��: Company Name
    Response.Write " .txtLoanDt.text        = """ & UNIDateClientFormat(E_f_ln_info(C_LOAN_DT)) & """      " & vbCr                                        '��: Plant Name
    Response.Write " .txtBasicLoanDt.text   = """ & UNIDateClientFormat(E_f_ln_info(C_LOAN_BASIC_DT)) & """ " & vbCr                                    '��: Currency Code
    Response.Write " .txtDeptCd.value       = """ & ConvSPChars(E_f_ln_info(C_DEPT_CD)) & """              " & vbCr                                            '��: Company Name
    Response.Write " .txtDeptNm.value       = """ & ConvSPChars(E_f_ln_info(C_DEPT_NM)) & """              " & vbCr                                        '��: Company FullName
    Response.Write " .txtLoanType.value     = """ & ConvSPChars(E_f_ln_info(C_LOAN_TYPE)) & """            " & vbCr
    Response.Write " .txtLoanTypeNm.value   = """ & ConvSPChars(E_f_ln_info(C_LOAN_TYPE_NM)) & """         " & vbCr
    Response.Write " .txtDueDt.text         = """ & UNIDateClientFormat(E_f_ln_info(C_DUE_DT)) & """      " & vbCr                                        '��: Plant Name
    Response.Write " .txt1stPrRdpDt.text    = """ & UNIDateClientFormat(E_f_ln_info(C_ST_PR_RDP_DT)) & """ " & vbCr
    Response.Write " .txtLoanAmt.text      = """ & UNINumClientFormat(E_f_ln_info(C_LOAN_AMT),				ggAmtOfMoney.DecPoint	,0)				& """" & vbCr    
    Response.Write " .txtLoanLocAmt.text   = """ & UNINumClientFormat(E_f_ln_info(C_LOAN_LOC_AMT),			ggAmtOfMoney.DecPoint	,0)				& """" & vbCr
    Response.Write " .txtBasRdpAmt.text    = """ & UNINumClientFormat(E_f_ln_info(C_BAS_RDP_AMT),			ggAmtOfMoney.DecPoint	,0)				& """" & vbCr    
    Response.Write " .txtBasRdpLocAmt.text = """ & UNINumClientFormat(E_f_ln_info(C_BAS_RDP_LOC_AMT),		ggAmtOfMoney.DecPoint	,0)				& """" & vbCr
    Response.Write " .txtBasIntPayAmt.text = """ & UNINumClientFormat(E_f_ln_info(C_BAS_INT_PAY_AMT),		ggAmtOfMoney.DecPoint	,0)				& """" & vbCr
    Response.Write " .txtBasIntPayLocAmt.text = """ & UNINumClientFormat(E_f_ln_info(C_BAS_INT_PAY_LOC_AMT), ggAmtOfMoney.DecPoint	,0)				& """" & vbCr
    Response.Write " .txtBpCd.value			= """ & ConvSPChars(E_f_ln_info(C_BP_CD)) & """					" & vbCr
    Response.Write " .txtBpNm.value			= """ & ConvSPChars(E_f_ln_info(C_BP_NM)) & """					" & vbCr
    Response.Write " .cboRdpClsFg.value     = """ & ConvSPChars(E_f_ln_info(C_RDP_CLS_FG)) & """           " & vbCr
    Response.Write " .txtGlNo.value         = """ & ConvSPChars(E_f_ln_info(C_GL_NO)) & """                 " & vbCr
    Response.Write " .cboPrRdpCond.value    = """ & ConvSPChars(E_f_ln_info(C_PR_RDP_COND)) & """          " & vbCr
    Response.Write " .txtPrRdpPerd.value    = """ & ConvSPChars(E_f_ln_info(C_PR_RDP_PERD)) & """          " & vbCr

    If ConvSPChars(E_f_ln_info(C_INT_VOTL)) = "X" Then															
		Response.Write " .Rb_IntVotl1.Checked	= """ & True & """											" & vbCr
		Else
		Response.Write " .Rb_IntVotl2.Checked    = """ & True & """										" & vbCr	
	End If
    
    Response.Write " .txt1stIntDueDt.text  = """ & UNIDateClientFormat(E_f_ln_info(C_ST_INT_DUE_DT)) & """        " & vbCr
    Response.Write " .txtIntRate.text      = """ & UNINumClientFormat(E_f_ln_info(C_LOAN_INT_RATE),		ggExchRate.DecPoint	,0)			& """" & vbCr        
    Response.Write " .txtIntPayPerd.value   = """ & ConvSPChars(E_f_ln_info(C_INT_PAY_PERD)) & """         " & vbCr    
    Response.Write " .cboIntPayStnd.value   = """ & E_f_ln_info(C_INT_PAY_STND) & """					   " & vbCr
    Response.Write " .cboIntBaseMthd.value		= """ & E_f_ln_info(C_INT_BASE_MTHD) & """                     " & vbCr
    
    Response.Write " .txtLoanDesc.value     = """ & ConvSPChars(E_f_ln_info(C_LOAN_DESC)) & """                         " & vbCr    

	Response.Write " .txtRdpAmt.text       = """ & UNINumClientFormat(E_f_ln_info(C_RDP_AMT),				ggAmtOfMoney.DecPoint	,0)				& """" & vbCr    
    Response.Write " .txtRdpLocAmt.text    = """ & UNINumClientFormat(E_f_ln_info(C_RDP_LOC_AMT),			ggAmtOfMoney.DecPoint	,0)				& """" & vbCr  

    Response.Write " .txtIntPayAmt.text    = """ & UNINumClientFormat(E_f_ln_info(C_INT_PAY_AMT),			ggAmtOfMoney.DecPoint	,0)				& """" & vbCr    
    Response.Write " .txtIntPayLocAmt.text = """ & UNINumClientFormat(E_f_ln_info(C_INT_PAY_LOC_AMT),		ggAmtOfMoney.DecPoint	,0)				& """" & vbCr         

    Response.Write " .txtLoanBalAmt.text   = """ & UNINumClientFormat(E_f_ln_info(C_LOAN_BAL_AMT),			ggAmtOfMoney.DecPoint	,0)				& """" & vbCr    
    Response.Write " .txtLoanBalLocAmt.text= """ & UNINumClientFormat(E_f_ln_info(C_LOAN_BAL_LOC_AMT),		ggAmtOfMoney.DecPoint	,0)				& """" & vbCr

    Response.Write " .cboLoanFg.value		= """ & ConvSPChars(E_f_ln_info(C_LOAN_FG)) & """         " & vbCr
    Response.Write " .txtDocCur.value		= """ & ConvSPChars(E_f_ln_info(C_DOC_CUR)) & """           " & vbCr
    Response.Write " .txtXchrate.text		= """ & UNINumClientFormat(E_f_ln_info(C_XCH_RATE),			ggExchRate.DecPoint	,0)				& """" & vbCr    
	Response.Write " .htxtLcNo.value			= """ & ConvSPChars(E_f_ln_info(C_LC_DOC_NO)) & """		   " & vbCr
	Response.Write " .htxtLoanPlcType.value		= """ & ConvSPChars(E_f_ln_info(C_LOAN_PLC_TYPE)) & """		   " & vbCr
	Response.Write " .hRdpSprdFg.value			= """ & ConvSPChars(E_f_ln_info(C_RDP_SPRD_FG)) & """		   " & vbCr
	Response.Write " .hClsRoFg.value			= """ & ConvSPChars(E_f_ln_info(C_CLS_RO_FG)) & """		   " & vbCr

    Response.Write " .htxtStIntPayAmt.value      = """ & UNINumClientFormat(E_f_ln_info(C_ST_ADV_INT_PAY_AMT),		ggAmtOfMoney.DecPoint	,0)				& """" & vbCr    
    Response.Write " .htxtStIntPayLocAmt.value   = """ & UNINumClientFormat(E_f_ln_info(C_ST_ADV_INT_PAY_LOC_AMT),	ggAmtOfMoney.DecPoint	,0)				& """" & vbCr    
    Response.Write " .txtPrRdpUnitAmt.value	= """ & UNINumClientFormat(E_f_ln_info(C_PR_RDP_UNIT_AMT),	ggAmtOfMoney.DecPoint	,0)				& """" & vbCr    
    Response.Write " .txtPrRdpUnitLocAmt.value = """ & UNINumClientFormat(E_f_ln_info(C_PR_RDP_UNIT_AMT),	ggAmtOfMoney.DecPoint	,0)     & """" & vbCr
	Response.Write " .txtTempGlNo.value			= """ & ConvSPChars(E_f_ln_info(C_TEMP_GL_NO)) & """		   " & vbCr
	Response.Write " .txtLoanAcctCd.value = """ & ConvSPChars(E_f_ln_info(C_LOAN_ACCT_CD)) & """         " & vbCr
	Response.Write " .txtLoanAcctNm.value = """ & ConvSPChars(E_f_ln_info(C_LOAN_ACCT_NM)) & """         " & vbCr
	Response.Write " .txtIntAcctCd.value = """ & ConvSPChars(E_f_ln_info(C_INT_ACCT_CD)) & """         " & vbCr
	Response.Write " .txtIntAcctNm.value = """ & ConvSPChars(E_f_ln_info(C_INT_ACCT_NM)) & """         " & vbCr
	Response.Write " .txtUserFld1.value = """ & ConvSPChars(E_f_ln_info(C_USER_FIELD1)) & """         " & vbCr
	Response.Write " .txtUserFld2.value = """ & ConvSPChars(E_f_ln_info(C_USER_FIELD2)) & """         " & vbCr
	Response.Write " .hOrgChangeId.value = """ & ConvSPChars(E_f_ln_info(C_ORG_CHANGE_ID)) & """         " & vbCr
    Response.Write " .txtStrFg.Value = """ & "B" & """														" & vbCr

    Response.Write "End with                " & vbCr
    Response.Write "Parent.DbQueryOk        " & vbCr
    Response.Write "</Script>               " & vbCr

End Sub
 
    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data To Db
'============================================================================================================
Sub SubBizSave()
	On Error Resume Next                                                             '��: Protect system from crashing
	Err.Clear        


	Dim iPAFG405CU

	Dim hOrgChangeId
	Dim I1_f_ln_info 
	Dim I2_a_temp_amt  
	Dim I3_b_biz_partner
	Dim I4_f_ln_info 
	Dim I5_b_bank  
	Dim I6_b_bank
	Dim I7_b_bank_acct 
	Dim I8_b_bank    
	Dim I9_b_bank_acct
	Dim I10_f_ln_repay_mean 
	Dim I11_b_acct_dept 
	Dim I12_b_currency
	Dim E1_b_auto_numbering 
	Dim E2_b_auto_numbering

	Dim lgIntFlgMode

	hOrgChangeId = Trim(Request("hOrgChangeId"))
	'Rollover�� ���, ���� ���Աݹ�ȣ 
	Const C_REF_LOAN_NO = 0 
	Redim I1_f_ln_info(C_REF_LOAN_NO) 
		I1_f_ln_info(C_REF_LOAN_NO) = ""
		
	'Rollover�� ���, �δ��� 
	Const C_AMT = 0
	Const C_LOC_AMT = 1
	Const BP_AMT = 2
	Const BP_LOC_AMT = 3
	Redim I2_a_temp_amt(BP_LOC_AMT) 

	I2_a_temp_amt(C_AMT) = 0
	I2_a_temp_amt(C_LOC_AMT) = 0
	I2_a_temp_amt(BP_AMT) = 0
	I2_a_temp_amt(BP_LOC_AMT) = 0

	'�ŷ�ó���Ա��� ���, bp_cd
	Const C_BP_CD = 0
	Redim I3_b_biz_partner(C_BP_CD) 

	I3_b_biz_partner(C_BP_CD) = Trim(Request("txtBpCd"))

	'���Ա� Main Data
	Const C_LOAN_NO = 0
	Const C_LOAN_NM = 1
	Const C_LOAN_FG = 2
	Const C_LOAN_DT = 3
	Const C_DOC_CUR = 4
	Const C_XCH_RATE = 5
	Const C_LOAN_AMT = 6
	Const C_LOAN_LOC_AMT = 7
	Const C_RDP_AMT = 8
	Const C_RDP_LOC_AMT = 9
	Const C_INT_PAY_AMT = 10
	Const C_INT_PAY_LOC_AMT = 11
	Const C_LOAN_TYPE = 12
	Const C_DUE_DT = 13
	Const C_RCPT_TYPE = 14
	Const C_LOAN_INT_RATE= 15
	Const C_LOAN_TERM = 16
	Const C_PR_RDP_TIMES = 17
	Const C_ST_PR_RDP_DT = 18
	Const C_ST_INT_DUE_DT = 19
	Const C_PR_RDP_PERD = 20
	Const C_PR_RDP_COND = 21
	Const C_INT_PAY_PERD = 22
	Const C_INT_PAY_PERD_BASE = 23
	Const C_INT_PAY_STND = 24
	Const C_INT_BASE_MTHD = 25
	Const C_INT_VOTL = 26
	Const C_DAY_MTHD = 27
	Const C_RDP_CLS_FG = 28
	Const C_LC_DOC_NO = 29
	Const C_PR_RDP_DT = 30
	Const C_INT_PAY_DT = 31
	Const C_ST_ADV_INT_PAY_AMT = 32
	Const C_ST_ADV_INT_PAY_LOC_AMT = 33
	Const C_LOAN_BASIC_FG = 34
	Const C_LOAN_PLC_TYPE = 35
	Const C_LOAN_BASIC_DT = 36
	Const C_RDP_SPRD_FG = 37
	Const C_LOAN_DESC = 38
	Const C_INSRT_DT = 39
	Const C_INSRT_USE_ID = 40
	Const C_UPDT_DT = 41
	Const C_UPDT_USE_ID = 42
	Const C_CLS_RO_FG = 43
	Const C_BAS_RDP_AMT = 44
	Const C_BAS_RDP_LOC_AMT = 45
	Const C_BAS_INT_PAY_AMT = 46
	Const C_BAS_INT_PAY_LOC_AMT = 47
	Const C_LOAN_BAL_AMT = 48
	Const C_LOAN_BAL_LOC_AMT = 49
	Const PR_RDP_UNIT_AMT = 50
	Const PR_RDP_UNIT_LOC_AMT = 51
	Const INTERNAL_CD = 52
	Const BP_CD = 53
    Const C_LOAN_ACCT_CD = 54		'���Աݰ����ڵ� 
    Const C_RCPT_ACCT_CD = 55		'�Ա����������ڵ� 
    Const C_INT_ACCT_CD = 56		'�������ް����ڵ� 
    Const C_CHARGE_ACCT_CD = 57		'�δ������ 
    Const C_PENALTY_ACCT_CD = 58
	Const C_USER_FIELD1 = 59
	Const C_USER_FIELD2 = 60
	Const C_REF_FG = 61	

	Redim I4_f_ln_info(C_REF_FG)

	I4_f_ln_info(C_LOAN_NO)			= UCase(Trim(Request("txtLoanNo")))
	I4_f_ln_info(C_LOAN_NM)			= Trim(Request("txtLoanNm"))
	I4_f_ln_info(C_LOAN_FG)			= UCase(Request("cboLoanFg"))
	I4_f_ln_info(C_LOAN_DT)			= UniConvDate(Request("txtLoanDt"))
	I4_f_ln_info(C_DOC_CUR)			= UCase(Request("txtDocCur"))
	I4_f_ln_info(C_XCH_RATE)		= UNIConvNum(Request("txtXchRate"),0)	
	I4_f_ln_info(C_LOAN_AMT)		= UNIConvNum(Request("txtLoanAmt"),0)	
	I4_f_ln_info(C_LOAN_LOC_AMT)	= UNIConvNum(Request("txtLoanLocAmt"),0)
	I4_f_ln_info(C_RDP_AMT)			= UNIConvNum(Request("txtRdpAmt"),0)	
	I4_f_ln_info(C_RDP_LOC_AMT)		=  UNIConvNum(Request("txtRdpLocAmt"),0)
	I4_f_ln_info(C_INT_PAY_AMT)		= UNIConvNum(Request("txtIntPayAmt"),0)	
	I4_f_ln_info(C_INT_PAY_LOC_AMT) = UNIConvNum(Request("txtIntPayLocAmt"),0)
	I4_f_ln_info(C_BAS_RDP_AMT)		= UNIConvNum(Request("txtBasRdpAmt"),0)
	I4_f_ln_info(C_BAS_RDP_LOC_AMT) = UNIConvNum(Request("txtBasRdpLocAmt"),0)
	I4_f_ln_info(C_BAS_INT_PAY_AMT) = UNIConvNum(Request("txtBasIntPayAmt"),0)
	I4_f_ln_info(C_BAS_INT_PAY_LOC_AMT) = UNIConvNum(Request("txtBasIntPayLocAmt"),0)
	I4_f_ln_info(C_LOAN_TYPE)		= UCase(Request("txtLoanType"))
	I4_f_ln_info(C_DUE_DT)			= UniConvDate(Request("txtDueDt"))
	I4_f_ln_info(C_RCPT_TYPE)		= UCase(Request("txtRcptType"))	
	I4_f_ln_info(C_LOAN_INT_RATE)	= UNIConvNum(Request("txtIntRate"),0)	
	I4_f_ln_info(C_LOAN_TERM)		= UNIConvNum(Request("txtLoanTerm"),0)
	I4_f_ln_info(C_PR_RDP_TIMES)	= ""
	If Trim(Request("cboPrRdpCond")) = "EQ" Then
		I4_f_ln_info(C_ST_PR_RDP_DT) = UNIConvDate(Request("txt1stPrRdpDt"))			
	Else 
		I4_f_ln_info(C_ST_PR_RDP_DT) = UniConvDate(Request("txtDueDt"))				
	End If	
				
	If Trim(Request("cboIntPayStnd")) = "DI" Then												
	    I4_f_ln_info(C_ST_INT_DUE_DT) =  UNIConvDate(Request("txt1stIntDueDt"))			
	ELse
		I4_f_ln_info(C_ST_INT_DUE_DT) =  UniConvDate(Request("txtLoanDt"))				
	End If 		
	I4_f_ln_info(C_PR_RDP_PERD)		= UNIConvNum(Request("txtPrRdpPerd"),0)
	I4_f_ln_info(C_PR_RDP_COND)		= Trim(Request("cboPrRdpCond"))
	I4_f_ln_info(C_INT_PAY_PERD)	= Request("txtIntPayPerd")

	If Trim(Request("cboIntBaseMthd")) = "12" Then
		I4_f_ln_info(C_INT_PAY_PERD_BASE) = "M"
	ElseIf Trim(Request("cboIntBaseMthd")) = "365" Or Trim(Request("cboIntBaseMthd")) = "360" Then
		I4_f_ln_info(C_INT_PAY_PERD_BASE) = "D"
	End If

	I4_f_ln_info(C_INT_PAY_STND)	= Trim(Request("cboIntPayStnd"))
	I4_f_ln_info(C_INT_BASE_MTHD)	= UNIConvNum(Request("cboIntBaseMthd"),0)	
	I4_f_ln_info(C_INT_VOTL)		= Trim(Request("Radio_IntVotl"))	
	I4_f_ln_info(C_DAY_MTHD)		= "YN"
	I4_f_ln_info(C_RDP_CLS_FG)		= Trim(Request("cboRdpClsFg"))
'		I4_f_ln_info(C_LC_DOC_NO)		= Trim(Request("txtLcNo"))
	I4_f_ln_info(C_PR_RDP_DT)		= 0
	I4_f_ln_info(C_INT_PAY_DT)		= 0
	I4_f_ln_info(C_ST_ADV_INT_PAY_AMT) = UNIConvNum(Request("txtStIntPayAmt"),0)	
	I4_f_ln_info(C_ST_ADV_INT_PAY_LOC_AMT) = UNIConvNum(Request("txtStIntPayLocAmt"),0)	
	I4_f_ln_info(C_LOAN_BASIC_FG)	= Trim(Request("txtLoanBasicFg"))	
	I4_f_ln_info(C_LOAN_PLC_TYPE)	= UCase(Request("htxtLoanPlcType"))
	I4_f_ln_info(C_LOAN_BASIC_DT)	= UniConvDate(Request("txtBasicLoanDt"))
	If UCase(Request("hRdpSprdFg")) = "" Then
		I4_f_ln_info(C_RDP_SPRD_FG) =  "N"											
	Else	
		I4_f_ln_info(C_RDP_SPRD_FG) = UCase(Request("hRdpSprdFg"))					
	End If	
	If UCase(Request("hClsRoFg")) = "" Then
		I4_f_ln_info(C_CLS_RO_FG)	=  "N"												
	Else 
		I4_f_ln_info(C_CLS_RO_FG)	=  UCase(Request("hClsRoFg"))						
	End If				
	I4_f_ln_info(C_LOAN_DESC)		= Trim(Request("txtLoanDesc"))
	I4_f_ln_info(C_INSRT_USE_ID)	= gUsrID
	I4_f_ln_info(C_UPDT_USE_ID)		= gUsrID 
	I4_f_ln_info(C_LOAN_ACCT_CD)	= UCase(Request("txtLoanAcctCd"))
	I4_f_ln_info(C_RCPT_ACCT_CD)	= ""
	I4_f_ln_info(C_INT_ACCT_CD)		= UCase(Request("txtIntAcctCd"))
	I4_f_ln_info(C_CHARGE_ACCT_CD)	= ""
	I4_f_ln_info(C_PENALTY_ACCT_CD)		= ""
	I4_f_ln_info(C_USER_FIELD1)		= Trim(Request("txtUserFld1"))
	I4_f_ln_info(C_USER_FIELD2)		= Trim(Request("txtUserFld2"))
			
	I4_f_ln_info(C_REF_FG)			= "ER"
			
	'������������ 
	Const C_LOAN_BANK_CD = 0
	Redim I5_b_bank(C_LOAN_BANK_CD)
		I5_b_bank(C_LOAN_BANK_CD) = ""

	'�Ա�����cd
	Const C_BANK_CD = 0
	Redim I6_b_bank(C_BANK_CD)
		I6_b_bank(C_BANK_CD) = Trim(Request("txtBankCd"))
		
	'�Աݰ��¹�ȣ 
	Const C_BANK_ACCT_NO = 0	
	Redim I7_b_bank_acct(C_BANK_ACCT_NO) 
		I7_b_bank_acct(C_BANK_ACCT_NO) = Trim(Request("txtBankAcct"))

	'Rollover�� ���, �δ��� �����������cd
	Const C_MEAN_BANK_CD = 0
	Redim I8_b_bank(C_MEAN_BANK_CD)
		I8_b_bank(C_MEAN_BANK_CD) = ""

	'Rollover�� ���, �δ��� ������� ���¹�ȣ	
	Const C_MEAN_BANK_ACCT_NO = 0	
	Redim I9_b_bank_acct(C_MEAN_BANK_ACCT_NO)
		I9_b_bank_acct(C_MEAN_BANK_ACCT_NO) = ""

	'Rollover�� ���, ������� type
	Const C_MEAN_TYPE = 0	
	Redim I10_f_ln_repay_mean(C_MEAN_TYPE)
		I10_f_ln_repay_mean(C_MEAN_TYPE) = UCase(Request("txtRcptType"))

	'�μ�����	
	Const C_CHG_ORG_ID = 0
	Const C_DEPT_CD = 1
	ReDim I11_b_acct_dept(C_DEPT_CD)
		I11_b_acct_dept(C_CHG_ORG_ID) = hOrgChangeId
		I11_b_acct_dept(C_DEPT_CD) = Trim(Request("txtDeptCd"))
		
	'�ڱ���ȭ 
	Const C_DOC_LOC_CUR = 0
	Redim I12_b_currency(C_DOC_LOC_CUR)
		I12_b_currency(C_DOC_LOC_CUR) = UCase(gCurrency)

	Const E_Loan_auto_no = 0
	Redim E1_b_auto_numbering(E_Loan_auto_no)

    Set iPAFG405CU = server.CreateObject("PAFG405.cFMngLnSvr")   

    If CheckSYSTEMError(Err,True) = True Then				
		Exit Sub
    End If          
	     
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '��: Read Operayion Mode (CREATE, UPDATE)
	        
    Select Case lgIntFlgMode   
		Case  OPMD_CMODE                                                             '�� : Create							
			  Call iPAFG405CU.F_MANAGE_LN_SVR(gStrGlobalCollection, "CREATE",	I1_f_ln_info,	I2_a_temp_amt,	I3_b_biz_partner,_
												I4_f_ln_info,	I5_b_bank, I6_b_bank,	I7_b_bank_acct,		I8_b_bank,_
												I9_b_bank_acct, I10_f_ln_repay_mean,	I11_b_acct_dept,	I12_b_currency, _
												E1_b_auto_numbering,	E2_b_auto_numbering)
			 strLoanNo = E1_b_auto_numbering(E_Loan_auto_no)
        Case  OPMD_UMODE           
			  Call iPAFG405CU.F_MANAGE_LN_SVR(gStrGlobalCollection, "UPDATE",	I1_f_ln_info,	I2_a_temp_amt,	I3_b_biz_partner,_
												I4_f_ln_info,	I5_b_bank, I6_b_bank,	I7_b_bank_acct,		I8_b_bank,_
												I9_b_bank_acct, I10_f_ln_repay_mean,	I11_b_acct_dept,	I12_b_currency, _
												E1_b_auto_numbering, E2_b_auto_numbering)
			 strLoanNo = E1_b_auto_numbering(E_Loan_auto_no)
		Case Else
				
    End Select

    If CheckSYSTEMError(Err,True) = True Then
		Set iPAFG405CU = nothing
		Exit Sub	
    End If
		 
    Set iPAFG405CU = nothing
	    
	Response.Write "<Script Language=vbscript>					" & vbCr
	Response.Write " parent.DbSaveOk(""" & strLoanNo	& """)	" & vbCr
    Response.Write "</Script>									" & vbCr    

End Sub
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : DELETE Data 
'============================================================================================================
Sub SubBizDelete()
	On Error Resume Next                                                             '��: Protect system from crashing
	Err.Clear        

	Dim iPAFG405D
	Dim I1_f_ln_info 
	Dim I2_f_ln_info

	'Rollover�� ���, ���� ���Աݹ�ȣ 
	Const C_REF_LOAN_NO = 0 
	Redim I1_f_ln_info(C_REF_LOAN_NO) 
		I1_f_ln_info(C_REF_LOAN_NO) = ""

	'�����ϴ� ���Աݹ�ȣ 
	Const C_LOAN_NO = 0 
	Redim I2_f_ln_info(C_LOAN_NO) 
		I2_f_ln_info(C_LOAN_NO) = Trim(Request("txtLoanNo"))
		
    If Trim(Request("txtLoanNo")) = ""  Then
		Call DisplayMsgBox("900002", vbInformation, "", "", I_MKSCRIPT)	'��ȸ�� ���� �ϼ���.
		Response.End 
	End If
		
	Set iPAFG405D = server.CreateObject("PAFG405.cFMngLnSvr")   
	    
	If CheckSYSTEMError(Err, True) = True Then					
	   Exit Sub
	End If    
			
	Call iPAFG405D.F_MANAGE_LN_SVR(gStrGlobalCollection,"DELETE",I1_f_ln_info, _
														,	,	I2_f_ln_info)
	    
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAFG405D = nothing
		Exit Sub
	End If
		 
	Set iPAFG405D = nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbDeleteOk          " & vbCr
	Response.Write "</Script>                   " & vbCr

End Sub


'============================================================================================================
' Name : SubPlanExec
' Desc : ��ȯ���� 
'============================================================================================================
Sub SubPlanExec()

    On Error Resume Next                                                                 '��: Protect system from crashing
    Err.Clear                                                                            '��: Clear Error status

	Dim PAFG400EXE                    				                            '�� : �Է�/������ ComProxy Dll ��� ����(as0031
	Dim EG1_export_group

    If Trim(Request("txtLoanNo")) = ""  Then
		Call DisplayMsgBox("900002", vbInformation, "", "", I_MKSCRIPT)	'��ȸ�� ���� �ϼ���.
		Response.End 
	End If

    Set PAFG400EXE = Server.CreateObject("PAFG400.cFMngLnPlnSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If

    Call PAFG400EXE.F_MANAGE_LN_PLAN_SVR(gStrGloBalCollection, Request("txtLoanNo"), UniConvDate(Request("txtDateFr")), _
								UniConvDate(Request("txtDateTo")), EG1_export_group)

    If CheckSYSTEMError(Err, True) = True Then
       Set PAFG400EXE = Nothing
       Exit Sub
    End If    

    Set PAFG400EXE = Nothing

	If IsEmpty(EG1_export_group) = False Then
	
		Call DisplayMsgBox("990000", vbOKOnly, "", "", I_MKSCRIPT)
		
	End If

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
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

%>