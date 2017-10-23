<%@ LANGUAGE=VBSCript %>
<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : a4106mb2.asp
'*  4. Program Name         : 채무반제(선금급) 저장 logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2000/03/30
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : You So Eun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☆) Means that "must change"
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
Dim lgIntFlgMode
    
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
	
	Dim iPAPG030	
	Dim iErrorPosition
	Dim iCommandSent
	Dim I1_a_acct_trans_type, I2_b_acct_dept, I3_b_bank
	Dim I4_b_bank_acct, I5_f_prpaym, I6_a_allc_paym
	Dim I7_b_biz_partner, I8_a_acct, I9_b_cost_center, I10_b_currency
	Dim IG1_import_group, IG2_import_group_dc, IG3_import_group_dc_dtl
	Dim E1_b_auto_numbering, E3_b_monthly_exchange_rate
	
	Dim LngMaxRow, LngMaxRow1, LngMaxRow3
	Dim lgIntFlgMode
	Dim LngRow, arrVal 
	
	Dim arrRowVal																	'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	
    Const A364_I2_org_change_id = 0
    Const A364_I2_dept_cd = 1

    Const A364_I6_paym_no = 0
    Const A364_I6_paym_dt = 1
    Const A364_I6_allc_type = 2
    Const A364_I6_paym_amt = 3
    Const A364_I6_paym_loc_amt = 4
    Const A364_I6_ref_no = 5
    Const A364_I6_diff_kind_cur = 6
    Const A364_I6_xch_rate = 7
    Const A364_I6_paym_type = 8
    Const A364_I6_note_no = 9
    Const A364_I6_diff_kind_cur_amt = 10
    Const A364_I6_diff_kind_cur_loc_amt = 11
    Const A364_I6_paym_desc = 12
    Const A364_I6_insrt_user_id = 13
    Const A364_I6_updt_user_id = 14
    Const A364_I6_dc_amt = 15
    Const A364_I6_dc_loc_amt = 16
    Const A364_I6_doc_cur = 17
    Const A364_I6_prpaym_no = 18
    Const A364_I6_insrt_dt = 19
    Const A364_I6_updt_dt = 20

'    Const A364_IG1_I1_ief_supplied_select_char = 0
'    Const A364_IG1_I2_a_open_ap_ap_no = 1
'    Const A364_IG1_I2_a_open_ap_ap_dt = 2
'    Const A364_IG1_I2_a_open_ap_doc_cur = 3
'    Const A364_IG1_I3_a_acct_acct_cd = 4
'    Const A364_IG1_I4_a_cls_ap_cls_dt = 5
'    Const A364_IG1_I4_a_cls_ap_doc_cur = 6
'    Const A364_IG1_I4_a_cls_ap_diff_kind_cur = 7
'    Const A364_IG1_I4_a_cls_ap_xch_rate = 8
'    Const A364_IG1_I4_a_cls_ap_cls_amt = 9
'    Const A364_IG1_I4_a_cls_ap_cls_loc_amt = 10
'    Const A364_IG1_I4_a_cls_ap_diff_kind_cur_amt = 11
'    Const A364_IG1_I4_a_cls_ap_diff_kind_cur_loc_amt = 12
'    Const A364_IG1_I4_a_cls_ap_cls_ap_desc = 13
'    Const A364_IG1_I4_a_cls_ap_dc_amt = 14
'    Const A364_IG1_I4_a_cls_ap_dc_loc_amt = 15
'    Const A364_IG1_I4_a_cls_ap_cls_type_fg = 16

'    Const A364_IG2_I1_ief_supplied_select_char = 0
'    Const A364_IG2_I2_a_acct_acct_cd = 1
'    Const A364_IG2_I3_a_paym_dc_seq = 2
'    Const A364_IG2_I3_a_paym_dc_dc_amt = 3
'    Const A364_IG2_I3_a_paym_dc_dc_loc_amt = 4
'    Const A364_IG2_I3_a_paym_dc_dc_desc = 5

'    Const A364_IG3_I1_ief_supplied_select_char = 0
'    Const A364_IG3_I2_a_paym_dc_seq = 1
'    Const A364_IG3_I3_a_ctrl_item_ctrl_cd = 2
'    Const A364_IG3_I4_a_paym_dc_dtl_dtl_seq = 3
'    Const A364_IG3_I4_a_paym_dc_dtl_ctrl_val = 4

	Dim I11_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
	Const A364_I11_a_data_auth_data_BizAreaCd = 0
	Const A364_I11_a_data_auth_data_internal_cd = 1
	Const A364_I11_a_data_auth_data_sub_internal_cd = 2
	Const A364_I11_a_data_auth_data_auth_usr_id = 3 

    On Error Resume Next																	'☜: Protect system from crashing
    Err.Clear																				'☜: Clear Error status                                                            

  	Redim I11_a_data_auth(3)
	I11_a_data_auth(A364_I11_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I11_a_data_auth(A364_I11_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I11_a_data_auth(A364_I11_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I11_a_data_auth(A364_I11_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
    
	LngMaxRow = CInt(Request("txtMaxRows"))													'☜: 최대 업데이트된 갯수 
	LngMaxRow1 = CInt(Request("txtMaxRows1"))												'☜: 최대 업데이트된 갯수 
	LngMaxRow3 = CInt(Request("txtMaxRows3"))												'☜: 최대 업데이트된 갯수 
	lgIntFlgMode = CInt(Request("txtFlgMode"))												'☜: 저장시 Create/Update 판별 
	
	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If
	
	redim I6_a_allc_paym(A364_I6_updt_dt)
	redim I2_b_acct_dept(A364_I2_dept_cd)

	'-----------------------
	'Data manipulate area
	'-----------------------												'⊙: Single 데이타 저장 
	If Trim(Request("txtPPNo")) = "" Then
		Call DisplayMsgBox("111129", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
    End If
    
    If Trim(Request("txtSpread")) = "" Then
		Call DisplayMsgBox("111310", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
    End If
    
	I1_a_acct_trans_type                  = "AP002"
	I10_b_currency                        = gCurrency
	I2_b_acct_dept(A364_I2_org_change_id) = GetGlobalInf("gChangeOrgId")
	I6_a_allc_paym(A364_I6_paym_no)       = Trim(Request("txtAllcNo"))
	I6_a_allc_paym(A364_I6_paym_dt)       = UNIConvDate(Request("txtAllcDt"))
	I6_a_allc_paym(A364_I6_allc_type)     = "P"
	I6_a_allc_paym(A364_I6_updt_user_id)  = Request("txtUpdtUserId")
	I2_b_acct_dept(A364_I2_dept_cd)       = Trim(Request("txtDeptCd"))
	I7_b_biz_partner                      = Trim(Request("txtBpCd"))
	I5_f_prpaym                           = Trim(Request("txtPPNo"))
	I6_a_allc_paym(A364_I6_doc_cur)       = Request("txtDocCur")
	I6_a_allc_paym(A364_I6_paym_amt)      = UNIConvNum(Request("txtClsAmt"),0)
	I6_a_allc_paym(A364_I6_paym_loc_amt)  = 0
	I6_a_allc_paym(A364_I6_paym_desc)     = Trim(Request("txtAllcDesc")) 
	
	Set iPAPG030 = Server.CreateObject ("PAPG030.cAMntPpAllcSvr")	
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
	
    Call iPAPG030.A_MAINT_PREPAYM_ALLC_SVR(gStrGlobalCollection,iCommandSent,I1_a_acct_trans_type,I2_b_acct_dept,I3_b_bank, _
									I4_b_bank_acct, I5_f_prpaym, I6_a_allc_paym, I7_b_biz_partner, I8_a_acct, I9_b_cost_center, _
										Request("txtSpread"), I10_b_currency, Request("txtSpread2"), Request("txtSpread3"), E1_b_auto_numbering, E3_b_monthly_exchange_rate,I11_a_data_auth)
 
    If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG030 = Nothing
		Exit Sub
	End If

    Set iPAPG030 = Nothing
                                                       
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
