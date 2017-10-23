<%
Option Explicit		
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : A404MB2
'*  4. Program Name         : PAYMENT 저장하는 P/G
'*  5. Program Desc         : PAYMENT 저장하는 P/G
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : CHANG SUNG HEE
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<%
'#########################################################################################################
'												1. Include
'##########################################################################################################
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
	
	Dim iPAPG020	
	Dim iErrorPosition
	Dim iCommandSent
	Dim I1_a_acct_trans_type, I2_b_acct_dept
	Dim IG1_import_group_dc_dtl, IG2_import_group_dc, IG3_import_group
	Dim I3_b_bank_acct, I4_b_bank, I5_a_allc_paym, I6_b_biz_partner
	Dim I7_a_acct, I8_b_currency
	Dim E1_b_auto_numbering, E3_b_monthly_exchange_rate
	
	Dim LngMaxRow, LngMaxRow1, LngMaxRow3
	Dim lgIntFlgMode
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

    Const A363_IG3_I1_ief_supplied_select_char = 0
    Const A363_IG3_I2_a_open_ap_ap_no = 1
    Const A363_IG3_I2_a_open_ap_ap_dt = 2
    Const A363_IG3_I3_a_acct_acct_cd = 3
    Const A363_IG3_I4_a_cls_ap_cls_dt = 4
    Const A363_IG3_I4_a_cls_ap_doc_cur = 5
    Const A363_IG3_I4_a_cls_ap_diff_kind_cur = 6
    Const A363_IG3_I4_a_cls_ap_xch_rate = 7
    Const A363_IG3_I4_a_cls_ap_cls_amt = 8
    Const A363_IG3_I4_a_cls_ap_cls_loc_amt = 9
    Const A363_IG3_I4_a_cls_ap_diff_kind_cur_amt = 10
    Const A363_IG3_I4_a_cls_ap_diff_kind_cur_loc_amt = 11
    Const A363_IG3_I4_a_cls_ap_cls_ap_desc = 12
    Const A363_IG3_I4_a_cls_ap_cls_ap_no = 13
    Const A363_IG3_I4_a_cls_ap_dc_amt = 14
    Const A363_IG3_I4_a_cls_ap_dc_loc_amt = 15
    Const A363_IG3_I4_a_cls_ap_cls_type_fg = 16

    Const A363_E3_multi_divide = 0
    Const A363_E3_std_rate = 1

	Dim I9_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
	Const A363_I9_a_data_auth_data_BizAreaCd = 0
	Const A363_I9_a_data_auth_data_internal_cd = 1
	Const A363_I9_a_data_auth_data_sub_internal_cd = 2
	Const A363_I9_a_data_auth_data_auth_usr_id = 3 
 
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear																			 '☜: Clear Error status          

  	Redim I9_a_data_auth(3)
	I9_a_data_auth(A363_I9_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I9_a_data_auth(A363_I9_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I9_a_data_auth(A363_I9_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I9_a_data_auth(A363_I9_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
    
	LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
	LngMaxRow1 = CInt(Request("txtMaxRows1"))										'☜: 최대 업데이트된 갯수 
	LngMaxRow3 = CInt(Request("txtMaxRows3"))											'☜: 최대 업데이트된 갯수 
	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 
	
	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If
	
	redim I5_a_allc_paym(A363_I5_prpaym_no)
	redim I2_b_acct_dept(A363_I2_dept_cd)
	
	'-----------------------
	'Data manipulate area
	'-----------------------												'⊙: Single 데이타 저장 

	I1_a_acct_trans_type = "AP001"
	I8_b_currency = gCurrency
	I5_a_allc_paym(A363_I5_paym_no)			= Trim(Request("txtAllcNo"))
	I5_a_allc_paym(A363_I5_paym_dt)			= UNIConvDate(Request("txtAllcDt"))
	I5_a_allc_paym(A363_I5_allc_type)		= "A"
	I5_a_allc_paym(A363_I5_doc_cur)			= UCase(Trim(Request("txtDocCur")))
	I5_a_allc_paym(A363_I5_xch_rate)		= UNIConvNum(Request("txtXchRate"),1)
	I4_b_bank = Trim(Request("txtBankCd"))
	I3_b_bank_acct = Trim(Request("txtBankAcct"))
	I5_a_allc_paym(A363_I5_paym_type)		= Request("txtInputType")
	I5_a_allc_paym(A363_I5_note_no)			= Trim(Request("txtCheckCD"))
	I5_a_allc_paym(A363_I5_paym_amt)		= UNIConvNum(Request("txtPaymAmt"),0)
	If Request("cbSetPaymLocAmt") = "Y" Then
		I5_a_allc_paym(A363_I5_paym_loc_amt)   = UNIConvNum(Request("txtPaymLocAmt"),0) 
	End If
	I5_a_allc_paym(A363_I5_dc_amt)			= UNIConvNum(Request("txtDcAmt"),0)
	I5_a_allc_paym(A363_I5_dc_loc_amt)		= UNIConvNum(Request("txtDcLocAmt"),0)
	I5_a_allc_paym(A363_I5_paym_desc)		= Trim(Request("txtPaymDesc"))
	I2_b_acct_dept(A363_I2_org_change_id)	= Trim(Request("hOrgChangeId"))
	I2_b_acct_dept(A363_I2_dept_cd)			= Trim(Request("txtDeptCd"))
	I6_b_biz_partner = Trim(Request("txtBpCd"))
	I7_a_acct = Trim(Request("txtAcctCd"))
	If Trim(Request("txtSpread")) = "" Then
		Call DisplayMsgBox("111110", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
    End If
    
	Set iPAPG020 = Server.CreateObject ("PAPG020.cAMntPayAllcSvr")	
	
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
	
    Call iPAPG020.A_MAINT_PAYM_ALLC_SVR(gStrGlobalCollection,iCommandSent,I1_a_acct_trans_type,I2_b_acct_dept,Request("txtSpread3"), _
									    Request("txtSpread2"), I3_b_bank_acct, I4_b_bank, I5_a_allc_paym, I6_b_biz_partner, I7_a_acct, _
										Request("txtSpread"), I8_b_currency, E1_b_auto_numbering, E3_b_monthly_exchange_rate,I9_a_data_auth)
	
    If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG020 = Nothing
		Exit Sub
	End If

    Set iPAPG020 = Nothing
                                                       
    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
	If Trim(E1_b_auto_numbering) <> "" then 
		Response.Write ".frm1.txtAllcNo.value = """ & ConvSPChars(E1_b_auto_numbering)    & """" & vbCr 
	End If
    Response.Write " .DBSaveOk(""" & ConvSPChars(E1_b_auto_numbering)  & """)" & vbCr
    Response.Write "End With "				 & vbCr	  
    Response.Write "</Script>"           
   
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
