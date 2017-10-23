
<%@ LANGUAGE=VBSCript%>
<%  
Option Explicit  

%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%
   
    Dim lgIntFlgMode
    
    On Error Resume Next															'☜: Protect system from crashing
    Err.Clear																		'☜: Clear Error status

    Call HideStatusWnd	
	Call LoadBasisGlobalInf()
    															'☜: Hide Processing message
   	lgIntFlgMode = Trim(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별   	

 '---------------------------------------Common-----------------------------------------------------------
	Select Case lgIntFlgMode
        Case CStr(OPMD_CMODE)	
             Call SubBizSaveMulti()												
        Case CStr(OPMD_UMODE)														'☜: Save,Update
             Call SubBizSaveMulti()													'☜: Multi  --> Save,Update,Delete
        Case CStr(UID_M0003)														'☜: Delete
            Call SubBizDelete()														'☜: Single --> Delete
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
    
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   		                                                                    
	
	Dim iPAPG020	
	Dim iErrorPosition
	Dim iCommandSent
	Dim I1_a_acct_trans_type, I2_b_acct_dept
	Dim I3_b_bank_acct, I4_b_bank, I5_a_allc_paym, I6_b_biz_partner
	Dim I7_a_acct, I8_b_currency, I9_batch_paym_no
	Dim I10_note_due_dt
	Dim E1_b_auto_numbering, E3_b_monthly_exchange_rate
	
'	Dim LngMaxRow, LngMaxRow1, LngMaxRow3
'	Dim lgIntFlgMode
	Dim LngRow, arrVal 
	
	Dim arrRowVal																	'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	
    Const A363_I2_org_change_id = 0
    Const A363_I2_dept_cd = 1

    Const A363_IG1_I1_ief_supplied_select_char = 0
    Const A363_IG1_I2_a_paym_dc_seq = 1
    Const A363_IG1_I3_a_ctrl_item_ctrl_cd = 2
    Const A363_IG1_I4_a_paym_dc_dtl_dtl_seq = 3
    Const A363_IG1_I4_a_paym_dc_dtl_ctrl_val = 4

    Const A363_IG2_I1_ief_supplied_select_char = 0
    Const A363_IG2_I2_a_acct_acct_cd = 1
    Const A363_IG2_I3_a_paym_dc_seq = 2
    Const A363_IG2_I3_a_paym_dc_dc_amt = 3
    Const A363_IG2_I3_a_paym_dc_dc_loc_amt = 4
    Const A363_IG2_I3_a_paym_dc_dc_desc = 5
    Const A363_I5_paym_no = 0
    Const A363_I5_paym_dt = 1
    Const A363_I5_allc_type = 2
    Const A363_I5_paym_amt = 3
    Const A363_I5_paym_loc_amt = 4
    Const A363_I5_ref_no = 5
    Const A363_I5_diff_kind_cur = 6
    Const A363_I5_xch_rate = 7
    Const A363_I5_paym_type = 8
    Const A363_I5_note_no = 9
    Const A363_I5_diff_kind_cur_amt = 10
    Const A363_I5_diff_kind_cur_loc_amt = 11
    Const A363_I5_paym_desc = 12
    Const A363_I5_insrt_user_id = 13
    Const A363_I5_updt_user_id = 14
    Const A363_I5_dc_amt = 15
    Const A363_I5_dc_loc_amt = 16
    Const A363_I5_doc_cur = 17
    Const A363_I5_prpaym_no = 18
		
	Dim I11_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
	Const A363_I11_a_data_auth_data_BizAreaCd = 0
	Const A363_I11_a_data_auth_data_internal_cd = 1
	Const A363_I11_a_data_auth_data_sub_internal_cd = 2
	Const A363_I11_a_data_auth_data_auth_usr_id = 3 
 
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear																			 '☜: Clear Error status          

  	Redim I11_a_data_auth(3)
	I11_a_data_auth(A363_I11_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I11_a_data_auth(A363_I11_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I11_a_data_auth(A363_I11_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I11_a_data_auth(A363_I11_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))                              

	If Cstr(lgIntFlgMode) = Cstr(OPMD_CMODE) Then
		iCommandSent = "CREATE"
	ElseIf Cstr(lgIntFlgMode) = Cstr(OPMD_UMODE) Then
		iCommandSent = "UPDATE"
	End If
    

	redim I5_a_allc_paym(A363_I5_prpaym_no)
	redim I2_b_acct_dept(A363_I2_dept_cd)

	'-----------------------
	'Data manipulate area
	'-----------------------												'⊙: Single 데이타 저장 

	I1_a_acct_trans_type = "AP001"

	I2_b_acct_dept(A363_I2_org_change_id) = Trim(Request("hOrgChangeId"))
	I2_b_acct_dept(A363_I2_dept_cd) = Trim(UCase(Request("txtDeptCd")))
	
	I3_b_bank_acct = Trim(Request("txtBankAcct"))
	
	If UCase(Trim(UCase(Request("txtInputType")))) = "CP"  Then
		I4_b_bank = Trim(Request("txtCardCoCd"))
	Else
		I4_b_bank = Trim(Request("txtBankCd"))		
	End If		

	I5_a_allc_paym(A363_I5_paym_type) = Trim(UCase(Request("txtInputType")))
	I5_a_allc_paym(A363_I5_note_no)   = Trim(Request("txtCheckCD"))
	I5_a_allc_paym(A363_I5_paym_no)   = ""
	I5_a_allc_paym(A363_I5_paym_dt)   = UNIConvDate(Request("txtAllcDt"))
	I5_a_allc_paym(A363_I5_allc_type) = ""
	I5_a_allc_paym(A363_I5_doc_cur)   = Trim(UCase(Request("txtDocCur")))
	I5_a_allc_paym(A363_I5_paym_desc) = Request("txtPaymDesc")	
	
	I6_b_biz_partner = ""
	
	I7_a_acct = Trim(Request("txtAcctCd"))
	
	I8_b_currency = gCurrency
	
	I9_batch_paym_no = Trim(Request("txtBatchAllcNo"))

	I10_note_due_dt = UNIConvDate(Request("txtNoteDueDt"))
	
	Set iPAPG020 = Server.CreateObject ("PAPG020.cAMntPayAllcSvr")	

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
    Call iPAPG020.A_MAINT_BATCH_PAYM_ALLC_SVR(gStrGlobalCollection, iCommandSent,I1_a_acct_trans_type,I2_b_acct_dept, I3_b_bank_acct, I4_b_bank, _ 
					I5_a_allc_paym, I6_b_biz_partner, I7_a_acct, Trim(Request("txtSpread")), I8_b_currency, I9_batch_paym_no, I10_note_due_dt, _ 
					E1_b_auto_numbering, E3_b_monthly_exchange_rate,I11_a_data_auth)
    
'    Call iPAPG020.A_MAINT_BATCH_PAYM_ALLC_SVR(gStrGlobalCollection, iCommandSent,I1_a_acct_trans_type,I2_b_acct_dept, I3_b_bank_acct, I4_b_bank, _ 
'					I5_a_allc_paym, I6_b_biz_partner, I7_a_acct, Trim(Request("txtSpread")), I8_b_currency, I9_batch_paym_no, _ 
'					E1_b_auto_numbering, E3_b_monthly_exchange_rate)
					
    If CheckSYSTEMError(Err,True) = True Then
       Set iPAPG020 = Nothing
       Exit Sub
	End If

    Set iPAPG020 = Nothing
                                                       
    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
  If Trim(E1_b_auto_numbering) <> "" then 
    Response.Write ".frm1.txtBatchAllcNo.value = """ & ConvSPChars(E1_b_auto_numbering)    & """" & vbCr 
  End If
    Response.Write " .DBSaveOk(""" & ConvSPChars(E1_b_auto_numbering)  & """)" & vbCr
    Response.Write "End With "				 & vbCr	  
    Response.Write "</Script>"           
    
End Sub    

'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Dim iPAPG020
    Dim iCommandSent
    Dim I9_batch_paym_no
    Dim I5_a_allc_paym
    
    Const A363_I5_paym_type = 8
	Const A363_I5_prpaym_no = 18

	Dim I11_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
	Const A363_I11_a_data_auth_data_BizAreaCd = 0
	Const A363_I11_a_data_auth_data_internal_cd = 1
	Const A363_I11_a_data_auth_data_sub_internal_cd = 2
	Const A363_I11_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I11_a_data_auth(3)
	I11_a_data_auth(A363_I11_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I11_a_data_auth(A363_I11_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I11_a_data_auth(A363_I11_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I11_a_data_auth(A363_I11_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))       
	
	Redim I5_a_allc_paym(A363_I5_prpaym_no)	    
    
    iCommandSent = "DELETE"

    I9_batch_paym_no = Trim(Request("txtBatchAllcNo"))
	I5_a_allc_paym(A363_I5_paym_type) = Trim(UCase(Request("txtInputType")))        
    
    Set iPAPG020 = Server.CreateObject ("PAPG020.cAMntPayAllcSvr")	
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

    Call iPAPG020.A_MAINT_BATCH_PAYM_ALLC_SVR(gStrGlobalCollection,iCommandSent,,, , ,I5_a_allc_paym , , , , , I9_batch_paym_no,,,,I11_a_data_auth)	
	
    If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG020 = Nothing
		Exit Sub
	End If

    Set iPAPG020 = Nothing
    
    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write " parent.DbDeleteOk()  " & vbCr  
    Response.Write "</Script>"  
        
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

