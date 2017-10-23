<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncServer.asp" -->

<%
     
    Dim lgOpModeCRUD

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD      = Request("txtMode")                                          '¢Ð: Read Operation Mode (CRUD)

	Select Case lgOpModeCRUD    
		Case CStr(UID_M0001)                                                         '¢Ð: Query
'			Call SubBizQuery()
			Call SubBizQueryMulti()
		Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
'			Call SubBizSave()
'			Call SubBizSaveMulti()
		Case CStr(UID_M0003)                                                         '¢Ð: Delete
'			Call SubBizDelete()
	End Select


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
	
	 '-------------------------------------------------------------------------------------------------
	 'EXPORTS View »ó¼ö 
	 '-------------------------------------------------------------------------------------------------
	 Const A300_E1_biz_area_cd		= 0								'View Name : export b_biz_area
	 Const A300_E1_biz_area_nm		= 1

	 'Group Name : export_group
	 Const A300_EG1_E1_biz_area_cd	= 0								'View Name : export_ar b_biz_area
	 Const A300_EG1_E1_biz_area_nm	= 1
	 Const A300_EG1_E2_dept_cd		= 2								'View Name : export_cls b_acct_dept
	 Const A300_EG1_E2_dept_nm		= 3
	 Const A300_EG1_E3_cls_dt		= 4								'View Name : export a_cls_ar
	 Const A300_EG1_E3_cls_amt		= 5
	 Const A300_EG1_E3_cls_loc_amt	= 6
	 Const A300_EG1_E3_dc_amt		= 7
	 Const A300_EG1_E3_dc_loc_amt	= 8
	 Const A300_EG1_E3_cls_ar_no	= 9
	 Const A300_EG1_E4_acct_cd		= 10							'View Name : export_cls_ar a_acct
	 Const A300_EG1_E4_acct_nm		= 11
	 Const A300_EG1_E5_ar_no		= 12							'View Name : export a_open_ar
	 Const A300_EG1_E5_ar_dt		= 13
	 Const A300_EG1_E5_ar_amt		= 14
	 Const A300_EG1_E5_ar_loc_amt	= 15
	 Const A300_EG1_E5_ar_due_dt	= 16
	 Const A300_EG1_E5_bal_amt		= 17
	 Const A300_EG1_E5_bal_loc_amt	= 18


	 Const A300_E3_allc_no			= 0								'View Name : export a_allc_rcpt
	 Const A300_E3_allc_dt			= 1
	 Const A300_E3_allc_type		= 2
	 Const A300_E3_ref_no			= 3
	 Const A300_E3_allc_amt			= 4
	 Const A300_E3_allc_loc_amt		= 5
	 Const A300_E3_dc_amt			= 6
	 Const A300_E3_dc_loc_amt		= 7

	    
	 Const A300_E4_bp_cd			= 0								'View Name : export b_biz_partner
	 Const A300_E4_bp_nm			= 1

	    
	 Const A300_E6_dept_cd			= 0								'View Name : export b_acct_dept
	 Const A300_E6_dept_nm			= 1

	    
	 Const A300_E7_rcpt_no			= 0								'View Name : export a_rcpt
	 Const A300_E7_rcpt_dt			= 1
	 Const A300_E7_rcpt_amt			= 2
	 Const A300_E7_rcpt_loc_amt		= 3
	 Const A300_E7_allc_amt			= 4
	 Const A300_E7_allc_loc_amt		= 5
	 Const A300_E7_bal_amt			= 6
	 Const A300_E7_bal_loc_amt		= 7


	 Const A300_E8_allc_dt			= 0								'View Name : export a_allc_rcpt_assn
	 Const A300_E8_allc_amt			= 1
	 Const A300_E8_allc_loc_amt		= 2
	 Const A300_E8_doc_cur			= 3
	 Const A300_E8_xch_rate			= 4

	'Group Name : export_group_hq
	 Const A300_EG2_E1_hq_seq		= 0								'View Name : export a_allc_rcpt_hq
	 Const A300_EG2_E1_ref_no		= 1
	 Const A300_EG2_E1_allc_amt		= 2
	 Const A300_EG2_E1_allc_loc_amt	= 3
	 Const A300_EG2_E1_dc_amt		= 4
	 Const A300_EG2_E1_dc_loc_amt	= 5
	 Const A300_EG2_E2_biz_area_cd	= 6								'View Name : export_hq b_biz_area
	 Const A300_EG2_E2_biz_area_nm	= 7
	 Const A300_EG2_E3_dept_cd		= 8								'View Name : export_hq b_acct_dept
	 Const A300_EG2_E3_dept_nm		= 9
	
	
	Dim objPADG020	
	Dim IstrArNo
	Dim IstrAllcNo
	Dim iStrData1
	Dim iStrData2
	Dim E1_b_biz_area
	Dim E2_a_open_ar_ar_no
	Dim EG1_export_group
	Dim E3_a_allc_rcpt
	Dim E4_b_biz_partner
	Dim E5_a_gl_gl_no
	Dim E6_b_acct_dept
	Dim E7_a_rcpt
	Dim E8_a_allc_rcpt_assn
	Dim EG2_export_group_hq
	Dim LngMaxRow
	Dim IntRows	
	Dim StrNextKey
	
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear          
	
	IstrAllcNo	= Request("txtAllcNo")
	IstrArNo	= Request("lgStrPrevKey")
	LngMaxRow	= Request("txtMaxRows")
	
	Set objPADG020 = CreateObject("PADG020.cALkUpAllcRcByHqSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If
    
	call objPADG020.A_LOOKUP_ALLC_RCPT_BY_HQ_SVR(gStrGloBalCollection,IstrArNo,IstrAllcNo,E1_b_biz_area, _
												 E2_a_open_ar_ar_no, EG1_export_group, E3_a_allc_rcpt, _
												 E4_b_biz_partner, E5_a_gl_gl_no, E6_b_acct_dept, _
												 E7_a_rcpt, E8_a_allc_rcpt_assn, EG2_export_group_hq)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set objPADG020 = Nothing
       Exit Sub
    End If    
    
	Set objPADG020 = Nothing
	
	If IsArray(EG1_export_group) Then
	
		For IntRows = 0 To Ubound(EG1_export_group)
	
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A300_EG1_E5_ar_no))
		    iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A300_EG1_E4_acct_cd))
		    iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A300_EG1_E4_acct_nm))
		    iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A300_EG1_E1_biz_area_cd))
		    iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A300_EG1_E1_biz_area_nm))
		    iStrData1 = iStrData1 & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows, A300_EG1_E5_ar_dt))
		    iStrData1 = iStrData1 & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows, A300_EG1_E5_ar_due_dt))
		    iStrData1 = iStrData1 & Chr(11) & UNINumClientFormat(EG1_export_group(IntRows, A300_EG1_E5_ar_amt), ggAmtOfMoney.DecPoint, 0)
		    iStrData1 = iStrData1 & Chr(11) & UNINumClientFormat(EG1_export_group(IntRows, A300_EG1_E5_bal_amt), ggAmtOfMoney.DecPoint, 0)
		    iStrData1 = iStrData1 & Chr(11) & UNINumClientFormat(EG1_export_group(IntRows, A300_EG1_E3_cls_amt), ggAmtOfMoney.DecPoint, 0) 
		    iStrData1 = iStrData1 & Chr(11) & UNINumClientFormat(EG1_export_group(IntRows, A300_EG1_E3_cls_loc_amt), ggAmtOfMoney.DecPoint, 0)
		    iStrData1 = iStrData1 & Chr(11) & LngMaxRow + IntRows + 1
			iStrData1 = iStrData1 & Chr(11) & Chr(12)
	
		Next
		
	End If	

	If ConvSPChars(E2_a_open_ar_ar_no) = ConvSPChars(EG1_export_group(IntRows - 1, A300_EG1_E5_ar_no)) Then
		StrNextKey = ""   
    Else
		StrNextKey = ConvSPChars(E2_a_open_ar_ar_no)
	End If		
	
	If IsArray(EG2_export_group_hq) Then
		For IntRows = 0 To Ubound(EG2_export_group_hq)

			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(EG2_export_group_hq(IntRows, A300_EG2_E2_biz_area_cd))
			iStrData2 = iStrData2 & Chr(11) & ""
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(EG2_export_group_hq(IntRows, A300_EG2_E2_biz_area_nm))
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(EG2_export_group_hq(IntRows, A300_EG2_E3_dept_cd))
			iStrData2 = iStrData2 & Chr(11) & ""
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(EG2_export_group_hq(IntRows, A300_EG2_E3_dept_nm))     
			iStrData2 = iStrData2 & Chr(11) & UNINumClientFormat(EG2_export_group_hq(IntRows, A300_EG2_E1_allc_amt), ggAmtOfMoney.DecPoint, 0)
			iStrData2 = iStrData2 & Chr(11) & UNINumClientFormat(EG2_export_group_hq(IntRows, A300_EG2_E1_allc_loc_amt), ggAmtOfMoney.DecPoint, 0)
			iStrData2 = iStrData2 & Chr(11) & LngMaxRow + IntRows + 1
			iStrData2 = iStrData2 & Chr(11) & Chr(12)                                      
		
		Next
	End If	
	
	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write " With Parent				" & vbCr
	
	If IsArray(E1_b_biz_area) then
		Response.Write "	.frm1.txtBizCd.Value		= """ & ConvSPChars(E1_b_biz_area(A300_E1_biz_area_cd))											& """" & vbCr
		Response.Write "	.frm1.txtBizNm.Value	    = """ & ConvSPChars(E1_b_biz_area(A300_E1_biz_area_nm))											& """" & vbCr
	End If
	
	If IsArray(E3_a_allc_rcpt) then
		Response.Write "	.frm1.txtAllcDt.text		= """ & UNIDateClientFormat(E3_a_allc_rcpt(A300_E3_allc_dt))									& """" & vbCr	
    End If
    
    If IsArray(E4_b_biz_partner) then
		Response.Write "	.frm1.txtBpCd.Value			= """ & ConvSPChars(E4_b_biz_partner(A300_E4_bp_cd))											& """" & vbCr
		Response.Write "	.frm1.txtBpNm.Value			= """ & ConvSPChars(E4_b_biz_partner(A300_E4_bp_nm))											& """" & vbCr   
    End If
    
    If Trim(E5_a_gl_gl_no) <> "" then
	    Response.Write "	.frm1.txtGlNo.value			= """ & ConvSPChars(E5_a_gl_gl_no)																& """" & vbCr
    End If
    
    If IsArray(E7_a_rcpt) then
		Response.Write "	.frm1.txtRcptDt.text		= """ & UNIDateClientFormat(E7_a_rcpt(A300_E7_rcpt_dt))											& """" & vbCr	    
		Response.Write "	.frm1.txtRcptNo.Value		= """ & ConvSPChars(E7_a_rcpt(A300_E7_rcpt_no))													& """" & vbCr    
		Response.Write "	.frm1.txtClsAmt.Text    	= """ & UNINumClientFormat(E7_a_rcpt(A300_E7_allc_amt), ggAmtOfMoney.DecPoint, 0)				& """" & vbCr
		Response.Write "	.frm1.txtClsLocAmt.Text		= """ & UNINumClientFormat(E7_a_rcpt(A300_E7_allc_loc_amt), ggAmtOfMoney.DecPoint, 0)			& """" & vbCr
		Response.Write "	.frm1.txtBalAmt.Text		= """ & UNINumClientFormat(E7_a_rcpt(A300_E7_bal_amt), ggAmtOfMoney.DecPoint, 0)				& """" & vbCr
		Response.Write "	.frm1.txtBalLocAmt.Text		= """ & UNINumClientFormat(E7_a_rcpt(A300_E7_bal_loc_amt), ggAmtOfMoney.DecPoint, 0)			& """" & vbCr
    End If
    
    If IsArray(E8_a_allc_rcpt_assn) then
		Response.Write "	.frm1.txtDocCur.value		= """ & ConvSPChars(E8_a_allc_rcpt_assn(A300_E8_doc_cur))										& """" & vbCr
		Response.Write "	.frm1.txtXchRate.Text		= """ & UNINumClientFormat(E8_a_allc_rcpt_assn(A300_E8_xch_rate), ggExchRate.DecPoint, 0)		& """" & vbCr	
	End If
	
	Response.Write " 	.ggoSpread.Source =       .frm1.vspdData	" & vbCr
	Response.Write " 	.ggoSpread.SSShowData """ & iStrData1  & """" & vbCr	
	
	Response.Write " 	.ggoSpread.Source =       .frm1.vspdData1	" & vbCr 
	Response.Write " 	.ggoSpread.SSShowData  """ & iStrData2 & """" & vbCr	
		
	Response.Write " 	.DbQueryOk				" & vbCr
	Response.Write " End With					" & vbCr
	Response.Write "</Script>					" & vbCr 
	
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next
	Err.Clear
	
	Response.Write "<Script Language=vbscript>	" & vbcr	
	Response.Write " With Parent				" & vbCr
	Response.Write " 	.DbSaveOk 				" & vbCr
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
