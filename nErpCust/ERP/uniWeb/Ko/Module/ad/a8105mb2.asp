<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncServer.asp" -->

<%
     
    

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

'	Select Case lgOpModeCRUD    
'		Case CStr(UID_M0001)                                                         '¢Ð: Query
'			Call SubBizQuery()
'			Call SubBizQueryMulti()
'		Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
'			Call SubBizSave()
			Call SubBizSaveMulti()
'		Case CStr(UID_M0003)                                                         '¢Ð: Delete
'			Call SubBizDelete()
'	End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write " With Parent				" & vbCr		
	Response.Write " 	.DbDeleteOK				" & vbCr
	Response.Write " End With					" & vbCr
	Response.Write "</Script>					" & vbCr 	
	
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear          
	                                                             
	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write " With Parent				" & vbCr
	Response.Write " 	.DbQueryOk				" & vbCr
	Response.Write " End With					" & vbCr
	Response.Write "</Script>					" & vbCr 
	
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	'-------------------------------------------------------------------------------------------------
	'IMPORTS View »ó¼ö 
	'-------------------------------------------------------------------------------------------------
    Const A357_I1_org_change_id				= 0							'View Name : import b_acct_dept
    Const A357_I1_dept_cd					= 1
    
    'Group Name : import_group_hq
    Const A357_IG1_I1_biz_area_cd			= 0							'View Name : import_hq b_biz_area
    Const A357_IG1_I2_org_change_id			= 1							'View Name : import_hq b_acct_dept
    Const A357_IG1_I2_dept_cd				= 2
    Const A357_IG1_I3_hq_seq				= 3							'View Name : import a_allc_paym_hq
    Const A357_IG1_I3_ref_no				= 4
    Const A357_IG1_I3_allc_amt				= 5
    Const A357_IG1_I3_allc_loc_amt			= 6


    Const A357_I6_paym_no					= 0							'View Name : import a_allc_paym
    Const A357_I6_paym_dt					= 1
    Const A357_I6_allc_type					= 2
    Const A357_I6_paym_amt					= 3
    Const A357_I6_paym_loc_amt				= 4
    Const A357_I6_ref_no					= 5
    Const A357_I6_diff_kind_cur				= 6
    Const A357_I6_xch_rate					= 7
    Const A357_I6_paym_type					= 8
    Const A357_I6_note_no					= 9
    Const A357_I6_diff_kind_cur_amt			= 10
    Const A357_I6_diff_kind_cur_loc_amt		= 11
    Const A357_I6_paym_desc					= 12
    Const A357_I6_insrt_user_id				= 13
    Const A357_I6_updt_user_id				= 14
    Const A357_I6_dc_amt					= 15
    Const A357_I6_dc_loc_amt				= 16
    Const A357_I6_doc_cur					= 17
    Const A357_I6_prpaym_no					= 18


    'Group Name : import_group
    Const A357_IG2_I1_select_char			= 0							'View Name : import ief_supplied
    Const A357_IG2_I2_ap_no					= 1							'View Name : import a_open_ap
    Const A357_IG2_I2_ap_dt					= 2
    Const A357_IG2_I3_acct_cd				= 3							'View Name : import_cls_sp a_acct
    Const A357_IG2_I4_cls_dt				= 4							'View Name : import a_cls_ap
    Const A357_IG2_I4_doc_cur				= 5
    Const A357_IG2_I4_diff_kind_cur			= 6
    Const A357_IG2_I4_xch_rate				= 7
    Const A357_IG2_I4_cls_amt				= 8
    Const A357_IG2_I4_cls_loc_amt			= 9
    Const A357_IG2_I4_diff_kind_cur_amt		= 10
    Const A357_IG2_I4_diff_kind_cur_loc_amt	= 11
    Const A357_IG2_I4_cls_ap_desc			= 12
    Const A357_IG2_I4_cls_ap_no				= 13
    Const A357_IG2_I4_dc_amt				= 14
    Const A357_IG2_I4_dc_loc_amt			= 15
    Const A357_IG2_I4_cls_type_fg			= 16


	Dim objPADG025
	Dim iCommandSent
	Dim I1_b_acct_dept
	Dim I2_b_biz_partner_bp_cd
	Dim I3_b_bank_acct_bank_acct_no
	Dim I4_b_bank_bank_cd
	Dim I5_a_acct_trans_type_trans_type
	Dim IG1_import_group_hq
	Dim I6_a_allc_paym
	Dim IG2_import_group
	Dim I7_b_currency_currency
	Dim E1_b_auto_numbering_auto_no
	Dim E3_b_monthly_exchange_rate_std_rate
	
	Dim lgIntFlgMode
	Dim LngMaxRow
	Dim LngMaxRow1
	Dim LngRow
	Dim arrRowVal
	Dim arrVal
	Dim strStatus
	
	
	On Error Resume Next
	Err.Clear
	
	
	LngMaxRow		= UNIConvNum(Request("txtMaxRows"),0)
	LngMaxRow1		= UNIConvNum(Request("txtMaxRows1"),0)
	lgIntFlgMode	= Cint(Request("txtFlgMode"))
	
	
	Redim I6_a_allc_paym(A357_I6_prpaym_no)
	Redim I1_b_acct_dept(A357_I1_dept_cd)

	
	
	I1_b_acct_dept(A357_I1_org_change_id)	= gChangeOrgId
	I1_b_acct_dept(A357_I1_dept_cd)			= Request("txtDeptCd")
	I2_b_biz_partner_bp_cd					= Request("txtBpCd")
	I3_b_bank_acct_bank_acct_no				= Request("txtBankAcct")
	I4_b_bank_bank_cd						= Request("txtBankCd")
	I5_a_acct_trans_type_trans_type			= "AP001"
	I6_a_allc_paym(A357_I6_paym_no)			= Request("txtAllcNo")
	I6_a_allc_paym(A357_I6_paym_dt)			= UNIConvDate(Request("txtAllcDt"))
	I6_a_allc_paym(A357_I6_allc_type)		= "H"
	I6_a_allc_paym(A357_I6_doc_cur)			= Request("txtDocCur")
	I6_a_allc_paym(A357_I6_xch_rate)		= UNIConvNum(Request("txtXchRate"),0)	
	I6_a_allc_paym(A357_I6_paym_type)		= Request("txtInputType")
	I6_a_allc_paym(A357_I6_note_no)			= Request("txtCheckCD")
	I6_a_allc_paym(A357_I6_paym_amt)		= UNIConvNum(Request("txtPaymAmt"),0)
	I6_a_allc_paym(A357_I6_paym_loc_amt)	= UNIConvNum(Request("txtPaymLocAmt"),0)
	I6_a_allc_paym(A357_I6_insrt_user_id)	= Request("txtUpdtUserId")
	I6_a_allc_paym(A357_I6_updt_user_id)	= Request("txtUpdtUserId")
	I7_b_currency_currency					= gCurrency
	

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	
	If Trim(Request("txtSpread")) <> "" Then
	
		Redim IG2_import_group(LngMaxRow - 1, A357_IG2_I4_cls_type_fg)
		
		arrRowVal = Split(Request("txtSpread"), gRowSep)

		For LngRow = 0 To LngMaxRow - 1
			    
			arrVal = Split(arrRowVal(LngRow), gColSep)
			strStatus = arrVal(0)
					            
			Select Case UCase(Trim(strStatus))

				Case "C", "U"
				
					IG2_import_group(LngRow, A357_IG2_I2_ap_no)			= arrVal(1)
					IG2_import_group(LngRow, A357_IG2_I3_acct_cd)		= arrVal(2)
					IG2_import_group(LngRow, A357_IG2_I2_ap_dt)			= UNIConvDate(ArrVal(3))
					IG2_import_group(LngRow, A357_IG2_I4_cls_dt)		= UNIConvDate(Request("txtAllcDt"))
					IG2_import_group(LngRow, A357_IG2_I4_doc_cur)		= arrVal(4)
					IG2_import_group(LngRow, A357_IG2_I4_cls_amt)		= UNIConvNum(arrVal(5),0)
					IG2_import_group(LngRow, A357_IG2_I4_cls_loc_amt)	= UNIConvNum(arrVal(6),0)

				Case "D"

			End Select

'?			If LngRow > 47 Then
'?				Call DisplayMsgBox(111131, , vbInformation, "", "", I_MKSCRIPT)
'?			End If

		Next

	End If
	

	If Trim(Request("txtSpread1")) <> "" Then

		Redim IG1_import_group_hq(LngMaxRow1 - 1, A357_IG1_I3_allc_loc_amt)
		
		arrRowVal = Split(Request("txtSpread1"), gRowSep)

		For LngRow = 0 To LngMaxRow1 - 1

			arrVal = Split(arrRowVal(LngRow), gColSep)
			strStatus = arrVal(0)
					            
			Select Case UCase(Trim(strStatus))

				Case "C", "U"

				IG1_import_group_hq(LngRow, A357_IG1_I3_hq_seq)			= UNIConvNum(arrVal(1),0)
				IG1_import_group_hq(LngRow, A357_IG1_I1_biz_area_cd)	= arrVal(2)
				IG1_import_group_hq(LngRow, A357_IG1_I2_dept_cd)		= arrVal(3)
				IG1_import_group_hq(LngRow, A357_IG1_I2_org_change_id)	= gChangeOrgId
				IG1_import_group_hq(LngRow, A357_IG1_I3_allc_amt)		= UNIConvNum(arrVal(4),0)
				IG1_import_group_hq(LngRow, A357_IG1_I3_allc_loc_amt)	= UNIConvNum(arrVal(5),0)

				Case "D"

			End Select		
						
'?			If LngRow > 47 Then
'?				Call DisplayMsgBox(111131, , vbInformation, "", "", I_MKSCRIPT)
'?			End If

		Next
		
	End If

	
	Set objPADG025 = CreateObject("PADG025.cAMntAllcPayHqSvr")
	
    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If

	Call objPADG025.A_MAINT_ALLC_PAYM_HQ_SVR (gStrGloBalCollection, iCommandSent, I1_b_acct_dept, I2_b_biz_partner_bp_cd, _
											  I3_b_bank_acct_bank_acct_no, I4_b_bank_bank_cd, I5_a_acct_trans_type_trans_type, _
											  IG1_import_group_hq, I6_a_allc_paym, IG2_import_group, I7_b_currency_currency, _
											  E1_b_auto_numbering_auto_no, E3_b_monthly_exchange_rate_std_rate)

	If CheckSYSTEMError(Err, True) = True Then
       Set objPADG025 = Nothing
		Exit Sub
    End If    										 
        
	Set objPADG025 = nothing

	
	Response.Write "<Script Language=vbscript>	" & vbcr	
	Response.Write " With Parent				" & vbCr	
	
	If ConvSPChars(E1_b_auto_numbering_auto_no) <> "" then
		Response.Write "	.frm1.txtAllcNo.value = """ & ConvSPChars(E1_b_auto_numbering_auto_no)	& """" & vbCr
	End If		
	
	Response.Write "	.DbSaveOk  """ & ConvSPChars(E1_b_auto_numbering_auto_no) 					& """" & vbCr
	Response.Write " End With					" & vbCr
	Response.Write "</Script>					" & vbCr
	
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next

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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

%>
