<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncServer.asp" -->

<%
     
    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

'	Select Case lgOpModeCRUD    
'		Case CStr(UID_M0001)                                                         '☜: Query
'			Call SubBizQuery()
'			Call SubBizQueryMulti()
'		Case CStr(UID_M0002)                                                         '☜: Save,Update
'			Call SubBizSave()
			Call SubBizSaveMulti()
'		Case CStr(UID_M0003)                                                         '☜: Delete
'			Call SubBizDelete()
'	End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
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
	
    On Error Resume Next                                                             '☜: Protect system from crashing
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
	'IMPORTS View 상수 
	'-------------------------------------------------------------------------------------------------
    'Group Name : import_group_ar
    Const A359_IG1_I1_select_char	= 0								'View Name : import ief_supplied
    Const A359_IG1_I2_acct_cd		= 1								'View Name : import_cls_ar a_acct
    Const A359_IG1_I3_ar_no			= 2								'View Name : import a_open_ar
    Const A359_IG1_I3_ar_dt			= 3
    Const A359_IG1_I4_cls_dt		= 4								'View Name : import a_cls_ar
    Const A359_IG1_I4_ar_due_dt		= 5
    Const A359_IG1_I4_doc_cur		= 6
    Const A359_IG1_I4_diff_kind_cur	= 7
    Const A359_IG1_I4_xch_rate		= 8
    Const A359_IG1_I4_cls_amt		= 9
    Const A359_IG1_I4_cls_loc_amt	= 10
    Const A359_IG1_I4_dc_amt		= 11
    Const A359_IG1_I4_dc_loc_amt	= 12
    Const A359_IG1_I4_cls_ar_desc	= 13
    
    
    Const A359_I2_org_change_id		= 0								'View Name : import b_acct_dept
    Const A359_I2_dept_cd			= 1


    Const A359_I3_allc_no			= 0								'View Name : import a_allc_rcpt
    Const A359_I3_allc_dt			= 1
    Const A359_I3_allc_type			= 2
    Const A359_I3_ref_no			= 3
    Const A359_I3_allc_amt			= 4
    Const A359_I3_allc_loc_amt		= 5
    Const A359_I3_dc_amt			= 6
    Const A359_I3_dc_loc_amt		= 7
    Const A359_I3_insrt_user_id		= 8
    Const A359_I3_insrt_dt			= 9
    Const A359_I3_updt_user_id		= 10
    Const A359_I3_updt_dt			= 11


    'Group Name : import_group_hq
    Const A359_IG2_I1_biz_area_cd	= 0								'View Name : import_hq b_biz_area
    Const A359_IG2_I2_org_change_id	= 1								'View Name : import_hq b_acct_dept
    Const A359_IG2_I2_dept_cd		= 2
    Const A359_IG2_I3_hq_seq		= 3								'View Name : import a_allc_rcpt_hq
    Const A359_IG2_I3_ref_no		= 4
    Const A359_IG2_I3_allc_amt		= 5
    Const A359_IG2_I3_allc_loc_amt	= 6


    Const A359_I4_rcpt_no			= 0								'View Name : import a_rcpt
    Const A359_I4_rcpt_dt			= 1


    Const A359_I5_allc_dt			= 0								'View Name : import a_allc_rcpt_assn
    Const A359_I5_doc_cur			= 1
    Const A359_I5_xch_rate			= 2
    Const A359_I5_allc_amt			= 3
    Const A359_I5_allc_loc_amt		= 4
    Const A359_I5_insrt_user_id		= 5
    Const A359_I5_updt_user_id		= 6


	Dim objPADG020
	Dim LngMaxRow
	Dim LngMaxRow1
	Dim lgIntFlgMode
	Dim arrRowVal
	Dim arrVal
	Dim LngRow
	Dim strStatus
	
	Dim iCommandSent
	Dim I1_b_currency_currency
	Dim I2_b_acct_dept
	Dim I3_a_allc_rcpt
	Dim I4_a_rcpt
	Dim I5_a_allc_rcpt_assn
	Dim I6_b_currency_currency
	Dim I7_a_acct_trans_type_trans_type
	Dim IG1_import_group_ar
	Dim IG2_import_group_hq
	Dim E1_b_auto_numbering_auto_no
		
	
	On Error Resume Next
	Err.Clear
	
	
	LngMaxRow		= UNIConvNum(Request("txtMaxRows"),0)										'최대 업데이트된 갯수 
	LngMaxRow1		= UNIConvNum(Request("txtMaxRows1"),0)										'최대 업데이트된 갯수 
	lgIntFlgMode	= Cint(Request("txtFlgMode"))										'저장시 Create/Update 판별 
    
	
	Redim I3_a_allc_rcpt(A359_I3_updt_dt)
	Redim I5_a_allc_rcpt_assn(A359_I5_updt_user_id)
	Redim I4_a_rcpt(A359_I4_rcpt_dt)
	Redim I2_b_acct_dept(A359_I2_dept_cd)
	
	
	I1_b_currency_currency						= gCurrency
	
	I2_b_acct_dept(A359_IG2_I2_org_change_id)	= gChangeOrgId
	
	I3_a_allc_rcpt(A359_I3_allc_no)				= Request("txtAllcNo")
	I3_a_allc_rcpt(A359_I3_allc_dt)				= UNIConvDate(Request("txtAllcDt"))
	I3_a_allc_rcpt(A359_I3_allc_type)			= "H"
	I3_a_allc_rcpt(A359_I3_allc_amt)			= UNIConvNum(Request("txtClsAmt"),0)
	I3_a_allc_rcpt(A359_I3_allc_loc_amt)		= UNIConvNum(Request("txtClsLocAmt"),0)
	I3_a_allc_rcpt(A359_I3_insrt_user_id)		= Request("txtUpdtUserId")
	I3_a_allc_rcpt(A359_I3_updt_user_id)		= Request("txtUpdtUserId")

	I4_a_rcpt(A359_I4_rcpt_no)					= Request("txtRcptNo")
	I4_a_rcpt(A359_I4_rcpt_dt)					= UNIConvDate(Request("txtRcptDt"))
	
	I5_a_allc_rcpt_assn(A359_I5_allc_dt)		= UNIConvDate(Request("txtAllcDt"))
	I5_a_allc_rcpt_assn(A359_I5_doc_cur)		= Request("txtDocCur")
	I5_a_allc_rcpt_assn(A359_I5_insrt_user_id)	= Request("txtUpdtUserId")
	I5_a_allc_rcpt_assn(A359_I5_updt_user_id)	= Request("txtUpdtUserId")
	
	I6_b_currency_currency						= gCurrency

	I7_a_acct_trans_type_trans_type				= "AR002"

	If lgIntFlgMode = OPMD_CMODE Then
	
		iCommandSent = "CREATE"
		
	ElseIf lgIntFlgMode = OPMD_UMODE Then
	
		iCommandSent = "UPDATE"
		
	End If
	

	If Trim(Request("txtSpread")) <> "" Then
		
		Redim IG1_import_group_ar(LngMaxRow - 1, A359_IG1_I4_cls_ar_desc)
		
		arrRowVal = Split(Request("txtSpread"), gRowSep)

		For LngRow = 0 to (LngMaxRow - 1)
		    
			arrVal = Split(arrRowVal(LngRow), gColSep)							
			strStatus = arrVal(0)														
			            
			Select Case UCase(Trim(strStatus))

				Case "C", "U"				
															
					IG1_import_group_ar(LngRow, A359_IG1_I3_ar_no)			= arrVal(1)
					IG1_import_group_ar(LngRow, A359_IG1_I2_acct_cd)		= arrVal(2)
					IG1_import_group_ar(LngRow, A359_IG1_I3_ar_dt)			= UNIConvDate(ArrVal(3))				
					IG1_import_group_ar(LngRow, A359_IG1_I4_cls_dt)			= UNIConvDate(Request("txtAllcDt"))				
					IG1_import_group_ar(LngRow, A359_IG1_I4_doc_cur)		= Request("txtDocCur")
					IG1_import_group_ar(LngRow, A359_IG1_I4_cls_amt)		= UNIConvNum(arrVal(4),0)
					IG1_import_group_ar(LngRow, A359_IG1_I4_cls_loc_amt)	= UNIConvNum(arrVal(5),0)
											
				Case "D"														
						
			End Select		
				
'??			If (LngRow + 1) > 48 Then
'??	            Call DisplayMsgBox(111131, , vbInformation, "", "", I_MKSCRIPT)			'⊙: you must release this line if you change msg into code
'??	        End If		

		Next
		
	End If

	If Trim(Request("txtSpread1")) <> "" Then
		
		Redim IG2_import_group_hq(LngMaxRow1 - 1, A359_IG2_I3_allc_loc_amt)
		
		arrRowVal = Split(Request("txtSpread1"), gRowSep)

		For LngRow = 0 to (LngMaxRow1 - 1)
		    
			arrVal = Split(arrRowVal(LngRow), gColSep)
			strStatus = arrVal(0)					
			            
			Select Case UCase(Trim(strStatus))

				Case "C", "U"
				
					IG2_import_group_hq(LngRow, A359_IG2_I3_hq_seq)			= UNIConvNum(arrVal(1),0)
					IG2_import_group_hq(LngRow, A359_IG2_I1_biz_area_cd)	= arrVal(2)
					IG2_import_group_hq(LngRow, A359_IG2_I2_dept_cd)		= arrVal(3)
					IG2_import_group_hq(LngRow, A359_IG2_I2_org_change_id)	= gChangeOrgId
					IG2_import_group_hq(LngRow, A359_IG2_I3_allc_amt)		= UNIConvNum(arrVal(4),0)
					IG2_import_group_hq(LngRow, A359_IG2_I3_allc_loc_amt)	= UNIConvNum(arrVal(5),0)
											
				Case "D"													
						
			End Select			
				
'??			If (LngRow + 1) > 48 Then
'??	            Call DisplayMsgBox(111131, , vbInformation, "", "", I_MKSCRIPT)			'⊙: you must release this line if you change msg into code
'??	        End If		

		Next
		
	End If	
	

	Set objPADG020 = CreateObject("PADG020.cAMntAllcRcHqSvr")
	
    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If


	E1_b_auto_numbering_auto_no = objPADG020.A_MAINT_ALLC_RCPT_HQ_SVR (gStrGloBalCollection, iCommandSent, IG1_import_group_ar, _
											I1_b_currency_currency , I2_b_acct_dept,I3_a_allc_rcpt ,IG2_import_group_hq ,I4_a_rcpt, _
											I5_a_allc_rcpt_assn,I6_b_currency_currency,I7_a_acct_trans_type_trans_type)											

	If CheckSYSTEMError(Err, True) = True Then
       Set objPADG020 = Nothing
		Exit Sub
    End If    										 
        
	Set objPADG020 = nothing

	
	Response.Write "<Script Language=vbscript>	" & vbcr	
	Response.Write " With Parent				" & vbCr	
	If ConvSPChars(E1_b_auto_numbering_auto_no) <> "" Then
		Response.Write " 	.frm1.txtAllcNo.value = """ & ConvSPChars(E1_b_auto_numbering_auto_no)	& """" & vbCr	
	END IF	
		Response.Write " 	.DbSaveOk """ & ConvSPChars(E1_b_auto_numbering_auto_no)				& """" & vbCr	
	Response.Write " End With					" & vbCr
	Response.Write "</Script>					" & vbCr
	
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

%>
