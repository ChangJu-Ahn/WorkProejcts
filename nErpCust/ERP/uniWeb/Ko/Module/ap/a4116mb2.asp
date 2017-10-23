
<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Account 
'*  2. Function Name        : 
'*  3. Program ID           : a4116mb2.adp
'*  4. Program Name         : (-)채권/출금반제 저장 Logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2000/03/30
'*  9. Modifier (First)     : YOU SO EUN
'* 10. Modifier (Last)      : YOU SO EUN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
														'☜ : ASP가 캐쉬되지 않도록 한다.
														'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%
   
On Error Resume Next															'☜: Protect system from crashing
Err.Clear																		'☜: Clear Error status

Call HideStatusWnd																'☜: Hide Processing message
Call LoadBasisGlobalInf()														

Call SubBizSaveMulti()															'☜: Multi  --> Save,Update,Delete

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
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
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
	
	Dim iPAPG080	
	Dim iErrorPosition
	Dim iCommandSent
	Dim I1_a_acct_trans_type, I2_b_acct_dept
	Dim I3_b_bank_acct, I4_b_bank, I5_a_allc_paym, I6_b_biz_partner
	Dim I7_a_acct, I8_b_currency
	Dim IG1_import_group_ar
	Dim E1_b_auto_numbering, E3_b_monthly_exchange_rate
	
	Dim LngMaxRow, LngMaxRow1, LngMaxRow3
	Dim lgIntFlgMode
	Dim LngRow, arrVal 
	
	Dim arrRowVal																	'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	
    Const A356_IG1_I1_select_char = 0    
    Const A356_IG1_I2_acct_cd = 1    
    Const A356_IG1_I3_ar_no = 2    
    Const A356_IG1_I3_ar_dt = 3
    Const A356_IG1_I4_cls_dt = 4    
    Const A356_IG1_I4_ar_due_dt = 5
    Const A356_IG1_I4_doc_cur = 6
    Const A356_IG1_I4_diff_kind_cur = 7
    Const A356_IG1_I4_xch_rate = 8
    Const A356_IG1_I4_cls_amt = 9
    Const A356_IG1_I4_cls_loc_amt = 10
    Const A356_IG1_I4_dc_amt = 11
    Const A356_IG1_I4_dc_loc_amt = 12
    Const A356_IG1_I4_cls_ar_desc = 13

    Const A356_I2_org_change_id = 0    
    Const A356_I2_dept_cd = 1

    Const A356_I5_paym_no = 0    
    Const A356_I5_paym_dt = 1
    Const A356_I5_allc_type = 2
    Const A356_I5_paym_amt = 3
    Const A356_I5_paym_loc_amt = 4
    Const A356_I5_ref_no = 5
    Const A356_I5_diff_kind_cur = 6
    Const A356_I5_xch_rate = 7
    Const A356_I5_paym_type = 8
    Const A356_I5_note_no = 9
    Const A356_I5_diff_kind_cur_amt = 10
    Const A356_I5_diff_kind_cur_loc_amt = 11
    Const A356_I5_paym_desc = 12
    Const A356_I5_insrt_user_id = 13
    Const A356_I5_updt_user_id = 14
    Const A356_I5_dc_amt = 15
    Const A356_I5_dc_loc_amt = 16
    Const A356_I5_doc_cur = 17
    Const A356_I5_prpaym_no = 18

	Dim I9_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
	Const A356_I9_a_data_auth_data_BizAreaCd = 0
	Const A356_I9_a_data_auth_data_internal_cd = 1
	Const A356_I9_a_data_auth_data_sub_internal_cd = 2
	Const A356_I9_a_data_auth_data_auth_usr_id = 3 
 
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear																			 '☜: Clear Error status          

  	Redim I9_a_data_auth(3)
	I9_a_data_auth(A356_I9_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I9_a_data_auth(A356_I9_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I9_a_data_auth(A356_I9_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I9_a_data_auth(A356_I9_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
    
	LngMaxRow = CInt(Request("txtMaxRows"))										'☜: 최대 업데이트된 갯수 
	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 
	
	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If
	
	Redim I5_a_allc_paym(A356_I5_prpaym_no)
	Redim I2_b_acct_dept(A356_I2_dept_cd)
	
	'-----------------------
	'Data manipulate area
	'-----------------------												'⊙: Single 데이타 저장 
	I1_a_acct_trans_type			      = "AP005"
	I8_b_currency                         = gCurrency
	I5_a_allc_paym(A356_I5_paym_no)       = Trim(Request("txtAllcNo"))
	I5_a_allc_paym(A356_I5_paym_dt)       = UNIConvDate(Request("txtAllcDt"))
	I5_a_allc_paym(A356_I5_allc_type)	  = "M"
	I5_a_allc_paym(A356_I5_doc_cur)		  = Request("txtDocCur")
	I4_b_bank							  = Trim(Request("txtBankCd"))
	I3_b_bank_acct						  = Trim(Request("txtBankAcct"))
	I5_a_allc_paym(A356_I5_paym_type)     = Request("txtInputType")
	I5_a_allc_paym(A356_I5_note_no)       = Trim(Request("txtCheckCD"))
	I5_a_allc_paym(A356_I5_paym_amt)	  = UNIConvNum(Request("txtPaymAmt"),0)
	I5_a_allc_paym(A356_I5_paym_loc_amt)  = UNIConvNum(Request("txtPaymLocAmt"),0) 
	I5_a_allc_paym(A356_I5_paym_desc)     = Trim(Request("txtAllcDesc"))  
	I2_b_acct_dept(A356_I2_org_change_id) = Trim(Request("hOrgChangeId"))
	I2_b_acct_dept(A356_I2_dept_cd)       = Trim(Request("txtDeptCd"))
	I6_b_biz_partner                      = Trim(Request("txtBpCd"))
	I7_a_acct                             = Trim(Request("txtAcctCd"))	
	
	If Trim(Request("txtSpread")) = "" Then
		Call DisplayMsgBox("112310", vbOKOnly, "", "", I_MKSCRIPT)
		Exit Sub
    End If
    		
	Set iPAPG080 = Server.CreateObject ("PAPG080.cAMntAllcPayByArSvr")	
	
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
	
    Call iPAPG080.A_MAINT_ALLC_PAYM_BY_AR_SVR(gStrGlobalCollection,iCommandSent,I1_a_acct_trans_type,Request("txtSpread"),I2_b_acct_dept, _
									I3_b_bank_acct, I4_b_bank, I5_a_allc_paym, I6_b_biz_partner, I7_a_acct, _
									I8_b_currency, E1_b_auto_numbering, E3_b_monthly_exchange_rate,I9_a_data_auth)
 
    If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG080 = Nothing
		Exit Sub
	End If

    Set iPAPG080 = Nothing
                                                       
    Response.Write "<Script Language=VBScript>												" & vbCr         
    Response.Write "With parent																" & vbCr	
	If Trim(E1_b_auto_numbering) <> "" then 
		Response.Write ".frm1.txtAllcNo.value = """ & ConvSPChars(E1_b_auto_numbering) & """" & vbCr 
	End If
    Response.Write " .DBSaveOk(""" & ConvSPChars(E1_b_auto_numbering)				  & """)" & vbCr
    Response.Write "End With																" & vbCr	  
    Response.Write "</Script>																" & vbCr
    
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
