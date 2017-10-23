<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 자금 
'*  3. Program ID           : f4235mb1
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
    
    Call HideStatusWnd   
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")                                                           '☜: Hide Processing message
	Dim lgOpModeCRUD
	gChangeOrgId = GetGlobalInf("gChangeOrgId")
    '---------------------------------------Common-----------------------------------------------------------
	lgErrorStatus	= "NO"
	lgErrorPos		= ""                                                           '☜: Set to space
	lgOpModeCRUD	= Request("txtMode")					'☜ : 현재 상태를 받음 

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
     Case CStr(UID_M0001)                                                         '☜: Query
          Call SubBizQuery()

     Case CStr(UID_M0002)       
          Call SubBizSave()

     Case CStr(UID_M0003)                                                         '☜: Delete
          Call SubBizDelete()
     Case "PAFG400"		'상환전개 
		  Call SubPlanExec()
End Select

Response.end

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status    

End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data To Db
'============================================================================================================
Sub SubBizSave()

	On Error Resume Next                                                             '☜: Protect system from crashing
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

	If Trim(Request("hOrgChangeId")) = "" Then
		hOrgChangeId = Trim(Request("hOrgChangeId"))
	Else
		hOrgChangeId = gChangeOrgId
	End IF
	'Rollover일 경우, 기존 차입금번호 
	Const C_REF_LOAN_NO = 0 
	Redim I1_f_ln_info(C_REF_LOAN_NO) 
		I1_f_ln_info(C_REF_LOAN_NO) = UCase(Trim(Request("txtLoanNo")))
		
	'Rollover일 경우, 부대비용 
	Const C_AMT = 0
	Const C_LOC_AMT = 1
	Const BP_AMT = 2
	Const BP_LOC_AMT = 3
	Redim I2_a_temp_amt(BP_LOC_AMT) 
		I2_a_temp_amt(C_AMT) = UNIConvNum(Request("txtChargeAmt"),0)
		I2_a_temp_amt(C_LOC_AMT) = UNIConvNum(Request("txtChargeLocAmt"),0)
		I2_a_temp_amt(BP_AMT) = UNIConvNum(Request("txtBPAmt"),0)
		I2_a_temp_amt(BP_LOC_AMT) = UNIConvNum(Request("txtBPLocAmt"),0)

	'거래처차입금일 경우, bp_cd
	Const C_BP_CD = 0
	Redim I3_b_biz_partner(C_BP_CD) 
		I3_b_biz_partner(C_BP_CD) = Trim(Request("txtBpRoCd"))

	'차입금 Main Data
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
	Const C_REF_NO = 29
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
    Const C_LOAN_ACCT_CD = 54		'차입금계정코드 
    Const C_RCPT_ACCT_CD = 55		'입금유형계정코드 
    Const C_INT_ACCT_CD = 56		'이자지급계정코드 
    Const C_CHARGE_ACCT_CD = 57		'부대비용계정 
    Const C_PENALTY_ACCT_CD = 58		'
	Const C_USER_FIELD1 = 59
	Const C_USER_FIELD2 = 60
	Const C_REF_FG = 61

	' -- 권한관리 
	Const A741_I3_a_data_auth_data_BizAreaCd = 0
	Const A741_I3_a_data_auth_data_internal_cd = 1
	Const A741_I3_a_data_auth_data_sub_internal_cd = 2
	Const A741_I3_a_data_auth_data_auth_usr_id = 3
	
	Dim I13_a_data_auth	'--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I13_a_data_auth(3)
	I13_a_data_auth(A741_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I13_a_data_auth(A741_I3_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I13_a_data_auth(A741_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I13_a_data_auth(A741_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	Redim I4_f_ln_info(C_REF_FG)

	I4_f_ln_info(C_LOAN_NO)			= UCase(Trim(Request("txtLoanRoNo")))
	I4_f_ln_info(C_LOAN_NM)			= Trim(Request("txtLoanRoNm"))
	I4_f_ln_info(C_LOAN_FG)			= UCase(Request("cboLoanFg"))
	I4_f_ln_info(C_LOAN_DT)			= UniConvDate(Request("txtLoanRoDt"))
	I4_f_ln_info(C_DOC_CUR)			= UCase(Request("txtDocCur"))
	I4_f_ln_info(C_XCH_RATE)		= UNIConvNum(Request("txtXchRate"),0)	
	I4_f_ln_info(C_LOAN_AMT)		= UNIConvNum(Request("txtLoanRoAmt"),0)	
	I4_f_ln_info(C_LOAN_LOC_AMT)	= UNIConvNum(Request("txtLoanRoLocAmt"),0)
	I4_f_ln_info(C_RDP_AMT)			= UNIConvNum(Request("txtTotPrRdpRoAmt"),0)
	I4_f_ln_info(C_RDP_LOC_AMT)		= UNIConvNum(Request("txtTotPrRdpRoLocAmt"),0)
	I4_f_ln_info(C_INT_PAY_AMT)		= UNIConvNum(Request("txtIntPayRoAmt"),0)	
	I4_f_ln_info(C_INT_PAY_LOC_AMT) = UNIConvNum(Request("txtIntPayLocAmt"),0)
	I4_f_ln_info(C_BAS_RDP_AMT)		= 0
	I4_f_ln_info(C_BAS_RDP_LOC_AMT) = 0
	I4_f_ln_info(C_BAS_INT_PAY_AMT) = 0
	I4_f_ln_info(C_BAS_INT_PAY_LOC_AMT) = 0
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
		I4_f_ln_info(C_ST_INT_DUE_DT) =  UniConvDate(Request("txtLoanRoDt"))
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
	I4_f_ln_info(C_DAY_MTHD)		= Trim(Request("Radio_IntStart")+Request("Radio_IntEnd"))	
	I4_f_ln_info(C_RDP_CLS_FG)		= Trim(Request("cboRdpClsFg"))
	I4_f_ln_info(C_PR_RDP_DT)		= 0
	I4_f_ln_info(C_INT_PAY_DT)		= 0
	I4_f_ln_info(C_ST_ADV_INT_PAY_AMT) = UNIConvNum(Request("txtStIntPayAmt"),0)	
	I4_f_ln_info(C_ST_ADV_INT_PAY_LOC_AMT) = UNIConvNum(Request("txtStIntPayLocAmt"),0)	
	I4_f_ln_info(C_LOAN_BASIC_FG)	= Trim(Request("txtLoanBasicFg"))	
	I4_f_ln_info(C_LOAN_PLC_TYPE)	= UCase(Request("htxtLoanPlcType"))
	I4_f_ln_info(C_LOAN_BASIC_DT)	= UniConvDate(Request("txtLoanDt"))
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
	I4_f_ln_info(C_RCPT_ACCT_CD)	= UCase(Request("txtRcptAcctCd"))
	I4_f_ln_info(C_INT_ACCT_CD)		= UCase(Request("txtIntAcctCd"))
	I4_f_ln_info(C_CHARGE_ACCT_CD)	= UCase(Request("txtChargeAcctCd"))
	I4_f_ln_info(C_PENALTY_ACCT_CD)	= UCase(Request("txtBPAcctCd"))
	I4_f_ln_info(C_USER_FIELD1)		= Trim(Request("txtUserFld1"))
	I4_f_ln_info(C_USER_FIELD2)		= Trim(Request("txtUserFld2"))
		
	'차입은행정보 
	Const C_LOAN_BANK_CD = 0
	Redim I5_b_bank(C_LOAN_BANK_CD)
		I5_b_bank(C_LOAN_BANK_CD) = ""

	'입금은행cd
	Const C_BANK_CD = 0
	Redim I6_b_bank(C_BANK_CD)
		I6_b_bank(C_BANK_CD) = Trim(Request("txtBankCd"))
		
	'입금계좌번호 
	Const C_BANK_ACCT_NO = 0	
	Redim I7_b_bank_acct(C_BANK_ACCT_NO) 
		I7_b_bank_acct(C_BANK_ACCT_NO) = Trim(Request("txtBankAcct"))

	'Rollover일 경우, 부대비용 출금유형은행cd
	Const C_MEAN_BANK_CD = 0
	Redim I8_b_bank(C_MEAN_BANK_CD)
		I8_b_bank(C_MEAN_BANK_CD) = Trim(Request("txtBankCd"))

	'Rollover일 경우, 부대비용 출금유형 계좌번호	
	Const C_MEAN_BANK_ACCT_NO = 0	
	Redim I9_b_bank_acct(C_MEAN_BANK_ACCT_NO)
		I9_b_bank_acct(C_MEAN_BANK_ACCT_NO) = Trim(Request("txtBankAcctNo"))

	'Rollover일 경우, 출금유형 type
	Const C_MEAN_TYPE = 0	
	Redim I10_f_ln_repay_mean(C_MEAN_TYPE)
		I10_f_ln_repay_mean(C_MEAN_TYPE) = UCase(Request("txtRcptType"))

	'부서정보	
	Const C_CHG_ORG_ID = 0
	Const C_DEPT_CD = 1
	ReDim I11_b_acct_dept(C_DEPT_CD)
		I11_b_acct_dept(C_CHG_ORG_ID) = hOrgChangeId
		I11_b_acct_dept(C_DEPT_CD) = Trim(Request("txtDeptCd"))
		
	'자국통화 
	Const C_DOC_LOC_CUR = 0
	Redim I12_b_currency(C_DOC_LOC_CUR)
		I12_b_currency(C_DOC_LOC_CUR) = UCase(gCurrency)
		
	Const E_Loan_auto_no = 0
	Redim E1_b_auto_numbering(E_Loan_auto_no)

    Set iPAFG405CU = server.CreateObject("PAFG405.cFMngLnSvr")   

    If CheckSYSTEMError(Err,True) = True Then				
		Exit Sub
    End If          
	     
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
	        
    Select Case lgIntFlgMode   
		Case  OPMD_CMODE                                                             '☜ : Create							
			  Call iPAFG405CU.F_MANAGE_LN_SVR(gStrGlobalCollection, "CREATE",	I1_f_ln_info,	I2_a_temp_amt,	I3_b_biz_partner,_
												I4_f_ln_info,	I5_b_bank, I6_b_bank,	I7_b_bank_acct,		I8_b_bank,_
												I9_b_bank_acct, I10_f_ln_repay_mean,	I11_b_acct_dept,	I12_b_currency, _
												E1_b_auto_numbering,	E2_b_auto_numbering, I13_a_data_auth)
			 strLoanNo = E1_b_auto_numbering(E_Loan_auto_no)
        Case  OPMD_UMODE           
			  Call iPAFG405CU.F_MANAGE_LN_SVR(gStrGlobalCollection, "UPDATE",	I1_f_ln_info,	I2_a_temp_amt,	I3_b_biz_partner,_
												I4_f_ln_info,	I5_b_bank, I6_b_bank,	I7_b_bank_acct,		I8_b_bank,_
												I9_b_bank_acct, I10_f_ln_repay_mean,	I11_b_acct_dept,	I12_b_currency, _
												E1_b_auto_numbering, E2_b_auto_numbering, I13_a_data_auth)
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

	Dim iPAFG405D
	Dim I1_f_ln_info 
	Dim I2_f_ln_info

	' -- 권한관리 
	Const A741_I3_a_data_auth_data_BizAreaCd = 0
	Const A741_I3_a_data_auth_data_internal_cd = 1
	Const A741_I3_a_data_auth_data_sub_internal_cd = 2
	Const A741_I3_a_data_auth_data_auth_usr_id = 3
	
	Dim I13_a_data_auth	'--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I13_a_data_auth(3)
	I13_a_data_auth(A741_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I13_a_data_auth(A741_I3_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I13_a_data_auth(A741_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I13_a_data_auth(A741_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear        

	'Rollover일 경우, 기존 차입금번호 
	Const C_REF_LOAN_NO = 0 
	Redim I1_f_ln_info(C_REF_LOAN_NO) 
		I1_f_ln_info(C_REF_LOAN_NO) = Trim(Request("txtLoanNo"))

	'삭제하는 차입금번호 
	Const C_LOAN_NO = 0 
	Redim I2_f_ln_info(C_LOAN_NO) 
		I2_f_ln_info(C_LOAN_NO) = Trim(Request("txtLoanRoNo"))
		
    If Trim(Request("txtLoanNo")) = ""  Then
		Call DisplayMsgBox("900002", vbInformation, "", "", I_MKSCRIPT)	'조회를 먼저 하세요.
		Response.End 
	End If


	Set iPAFG405D = server.CreateObject("PAFG405.cFMngLnSvr")   
	    
	If CheckSYSTEMError(Err, True) = True Then					
	   Exit Sub
	End If    
			
	Call iPAFG405D.F_MANAGE_LN_SVR(gStrGlobalCollection,"DELETE",I1_f_ln_info, _
														,	,	I2_f_ln_info, , , , , , , , , , I13_a_data_auth)
	    
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
' Desc : 상환전개 
'============================================================================================================
Sub SubPlanExec()

    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	Dim PAFG400EXE                    				                            '☆ : 입력/수정용 ComProxy Dll 사용 변수(as0031
	Dim EG1_export_group

    If Request("txtLoanNo") = ""  Then
		Call DisplayMsgBox("900002", vbInformation, "", "", I_MKSCRIPT)	'조회를 먼저 하세요.
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
%>

