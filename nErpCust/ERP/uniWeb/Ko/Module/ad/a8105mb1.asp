<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncServer.asp" -->

<%
     
    Dim lgOpModeCRUD

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD      = Request("txtMode")                                          '☜: Read Operation Mode (CRUD)

	Select Case lgOpModeCRUD    
		Case CStr(UID_M0001)                                                         '☜: Query
'			Call SubBizQuery()
			Call SubBizQueryMulti()
		Case CStr(UID_M0002)                                                         '☜: Save,Update
'			Call SubBizSave()
'			Call SubBizSaveMulti()
		Case CStr(UID_M0003)                                                         '☜: Delete
'			Call SubBizDelete()
	End Select


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
	
	 '-------------------------------------------------------------------------------------------------
	 'EXPORTS View 상수 
	 '-------------------------------------------------------------------------------------------------
    Const A291_E2_biz_area_cd		= 0								'  View Name : export b_biz_area
    Const A291_E2_biz_area_nm		= 1


    Const A291_E3_paym_no			= 0								'  View Name : export a_allc_paym
    Const A291_E3_paym_dt			= 1
    Const A291_E3_allc_type			= 2
    Const A291_E3_paym_amt			= 3
    Const A291_E3_paym_loc_amt		= 4
    Const A291_E3_ref_no			= 5
    Const A291_E3_diff_kind_cur		= 6
    Const A291_E3_xch_rate			= 7
    Const A291_E3_paym_type			= 8
    Const A291_E3_note_no			= 9
    Const A291_E3_paym_desc			= 10
    Const A291_E3_doc_cur			= 11


    Const A291_E4_bank_cd			= 0								'  View Name : export b_bank
    Const A291_E4_bank_nm			= 1


    Const A291_E5_acct_cd			= 0								'  View Name : export a_acct
    Const A291_E5_acct_nm			= 1


    Const A291_E6_bp_cd				= 0								'  View Name : export b_biz_partner
    Const A291_E6_bp_nm				= 1

    
    '  Group Name : export_group
    Const A291_EG1_E1_biz_area_cd	= 0								'  View Name : export_ap b_biz_area
    Const A291_EG1_E1_biz_area_nm	= 1
    Const A291_EG1_E2_dept_cd		= 2								'  View Name : export_cls_ap b_acct_dept
    Const A291_EG1_E2_dept_nm		= 3
    Const A291_EG1_E3_bp_cd			= 4								'  View Name : export_cls_ap b_biz_partner
    Const A291_EG1_E3_bp_nm			= 5
    Const A291_EG1_E4_acct_cd		= 6								'  View Name : export_cls_ap a_acct
    Const A291_EG1_E4_acct_nm		= 7
    Const A291_EG1_E5_cls_dt		= 8								'  View Name : export a_cls_ap
    Const A291_EG1_E5_doc_cur		= 9
    Const A291_EG1_E5_diff_kind_cur	= 10
    Const A291_EG1_E5_xch_rate		= 11
    Const A291_EG1_E5_cls_amt		= 12
    Const A291_EG1_E5_cls_loc_amt	= 13
    Const A291_EG1_E6_ap_no			= 14							'  View Name : export a_open_ap
    Const A291_EG1_E6_ap_dt			= 15
    Const A291_EG1_E6_doc_cur		= 16
    Const A291_EG1_E6_xch_rate		= 17
    Const A291_EG1_E6_ap_due_dt		= 18
    Const A291_EG1_E6_ap_amt		= 19
    Const A291_EG1_E6_ap_loc_amt	= 20
    Const A291_EG1_E6_bal_amt		= 21
    Const A291_EG1_E6_bal_loc_amt	= 22
    
    
    Const A291_E9_dept_cd			= 0								'  View Name : export b_acct_dept
    Const A291_E9_dept_nm			= 1


    '  Group Name : export_group_hq
    Const A291_EG2_E1_biz_area_cd	= 0								'  View Name : export_hq b_biz_area
    Const A291_EG2_E1_biz_area_nm	= 1
    Const A291_EG2_E2_dept_cd		= 2								'  View Name : exort_hq b_acct_dept
    Const A291_EG2_E2_dept_nm		= 3
    Const A291_EG2_E3_hq_seq		= 4								'  View Name : export a_allc_paym_hq
    Const A291_EG2_E3_allc_amt		= 5
    Const A291_EG2_E3_allc_loc_amt	= 6
    
    
	Dim objPADG025	
	Dim I1_a_allc_paym_paym_no
	Dim I2_a_open_ap_ap_no
	Dim I3_a_paym_dc_seq
	Dim I4_a_allc_paym_hq_hq_seq
	Dim E1_b_minor_minor_nm
	Dim E2_b_biz_area
	Dim E3_a_allc_paym
	Dim E4_b_bank
	Dim E5_a_acct
	Dim E6_b_biz_partner
	Dim E7_a_gl_gl_no
	Dim E8_a_open_ap_ap_no
	Dim EG1_export_group
	Dim E9_b_acct_dept
	Dim E10_b_bank_acct_bank_acct_no
	Dim EG2_export_group_hq
	Dim E11_a_allc_paym_hq_hq_seq
	Dim lgStrPrevKey
	Dim StrNextKey	
	Dim IntRows
	Dim LngMaxRow
	Dim LngMaxRow1
	Dim strData
	Dim strData1
	

	On Error Resume Next
	Err.Clear
	
		
	lgStrPrevKey			= Request("lgStrPrevKey")
	LngMaxRow				= UNIConvNum(Request("txtMaxRows"),0)
	LngMaxRow1				= UNIConvNum(Request("txtMaxRows1"),0)
	
	I1_a_allc_paym_paym_no	= Request("txtAllcNo")
	I2_a_open_ap_ap_no		= lgStrPrevKey
	
	Set objPADG025 = CreateObject("PADG025.cALkUpAllcPayByHqSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If
    
    
	call objPADG025.A_LOOKUP_ALLC_PAYM_BY_HQ_SVR (gStrGloBalCollection, I1_a_allc_paym_paym_no, I2_a_open_ap_ap_no, _
												 I3_a_paym_dc_seq, I4_a_allc_paym_hq_hq_seq, E1_b_minor_minor_nm, _
												 E2_b_biz_area, E3_a_allc_paym, E4_b_bank, E5_a_acct, E6_b_biz_partner, _
												 E7_a_gl_gl_no, E8_a_open_ap_ap_no, EG1_export_group, E9_b_acct_dept, _
												 E10_b_bank_acct_bank_acct_no, EG2_export_group_hq, E11_a_allc_paym_hq_hq_seq)
	
	If CheckSYSTEMError(Err, True) = True Then
       Set objPADG025 = Nothing
       Exit Sub
    End If
    
	Set objPADG025 = Nothing


	If IsArray(EG1_export_group) then

		If ConvSPChars(E8_a_open_ap_ap_no) = ConvSPChars(EG1_export_group(UBound(EG1_export_group),A291_EG1_E6_ap_no)) Then
			StrNextKey = ""   
		Else
			StrNextKey = ConvSPChars(E8_a_open_ap_ap_no)
		End If	
	
		For IntRows = 0 to UBound(EG1_export_group)
		
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A291_EG1_E6_ap_no))
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A291_EG1_E4_acct_cd))
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A291_EG1_E4_acct_nm))
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A291_EG1_E1_biz_area_cd))
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A291_EG1_E1_biz_area_nm))
			strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows, A291_EG1_E6_ap_dt))
			strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(IntRows, A291_EG1_E6_ap_due_dt))
			strData = strData & Chr(11) & ConvSPChars(EG1_export_group(IntRows, A291_EG1_E6_doc_cur))
			strData = strData & Chr(11) & UNINumClientFormat(EG1_export_group(IntRows, A291_EG1_E6_ap_amt), ggAmtOfMoney.DecPoint, 0)
			strData = strData & Chr(11) & UNINumClientFormat(EG1_export_group(IntRows, A291_EG1_E6_bal_amt), ggAmtOfMoney.DecPoint, 0)
			strData = strData & Chr(11) & UNINumClientFormat(EG1_export_group(IntRows, A291_EG1_E5_cls_amt), ggAmtOfMoney.DecPoint, 0)
			strData = strData & Chr(11) & UNINumClientFormat(EG1_export_group(IntRows, A291_EG1_E5_cls_loc_amt), ggAmtOfMoney.DecPoint, 0)
			strData = strData & Chr(11) & LngMaxRow + IntRows + 1
			strData = strData & Chr(11) & Chr(12)
		
		Next
		
	Else	
		Call ServerMesgBox("채권 상세 정보가 입력되어 있지 않습니다!", vbInformation, I_MKSCRIPT)
	End If

	
	If IsArray(EG2_export_group_hq) then
	
		For IntRows = 0 To UBound(EG2_export_group_hq)

			strData1 = strData1 & Chr(11) & ConvSPChars(EG2_export_group_hq(intRows, A291_EG2_E1_biz_area_cd))
		    strData1 = strData1 & Chr(11) & ""
		    strData1 = strData1 & Chr(11) & ConvSPChars(EG2_export_group_hq(intRows, A291_EG2_E1_biz_area_nm))
		    strData1 = strData1 & Chr(11) & ConvSPChars(EG2_export_group_hq(intRows, A291_EG2_E2_dept_cd))
		    strData1 = strData1 & Chr(11) & ""
		    strData1 = strData1 & Chr(11) & ConvSPChars(EG2_export_group_hq(intRows, A291_EG2_E2_dept_nm))
		    strData1 = strData1 & Chr(11) & UNINumClientFormat(EG2_export_group_hq(intRows, A291_EG2_E3_allc_amt), ggAmtOfMoney.DecPoint, 0)
		    strData1 = strData1 & Chr(11) & UNINumClientFormat(EG2_export_group_hq(intRows, A291_EG2_E3_allc_loc_amt), ggAmtOfMoney.DecPoint, 0)
		    strData1 = strData1 & Chr(11) & LngMaxRow1 + IntRows +1
			strData1 = strData1 & Chr(11) & Chr(12)

		Next
		
	End If
	

	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write " With Parent				" & vbCr

	If IsArray(E3_a_allc_paym) then
		Response.Write "	.frm1.txtAllcDt.text		= """ & UNIDateClientFormat(E3_a_allc_paym(A291_E3_paym_dt))	& """" & vbCr	
		Response.Write "	.frm1.txtInputType.Value	= """ & ConvSPChars(E3_a_allc_paym(A291_E3_paym_type))			& """" & vbCr
		Response.Write "	.frm1.txtInputTypeNm.Value	= """ & ConvSPChars(E1_b_minor_minor_nm)						& """" & vbCr
		Response.Write "	.frm1.txtCheckCd.Value		= """ & ConvSPChars(E3_a_allc_paym(A291_E3_note_no))			& """" & vbCr
		Response.Write "	.frm1.txtDocCur.value		= """ & ConvSPChars(E3_a_allc_paym(A291_E3_doc_cur))			& """" & vbCr
		Response.Write "	.frm1.txtXchRate.Text		= """ & UNINumClientFormat(E3_a_allc_paym(A291_E3_xch_rate), ggExchRate.DecPoint, 0)		& """" & vbCr
		Response.Write "	.frm1.txtPaymAmt.Text		= """ & UNINumClientFormat(E3_a_allc_paym(A291_E3_paym_amt), ggAmtOfMoney.DecPoint, 0)		& """" & vbCr
		Response.Write "	.frm1.txtPaymLocAmt.Text	= """ & UNINumClientFormat(E3_a_allc_paym(A291_E3_paym_loc_amt), ggAmtOfMoney.DecPoint, 0)	& """" & vbCr
	End If
	
	If IsArray(E9_b_acct_dept) then
		Response.Write "	.frm1.txtDeptCd.Value		= """ & ConvSPChars(E9_b_acct_dept(A291_E9_dept_cd))	& """" & vbCr
		Response.Write "	.frm1.txtDeptNm.Value	    = """ & ConvSPChars(E9_b_acct_dept(A291_E9_dept_nm))		& """" & vbCr
	End If
    
    If IsArray(E4_b_bank) then
		Response.Write "	.frm1.txtBankCd.Value		= """ & ConvSPChars(E4_b_bank(A291_E4_bank_cd))			& """" & vbCr
		Response.Write "	.frm1.txtBankNm.Value	    = """ & ConvSPChars(E4_b_bank(A291_E4_bank_nm))			& """" & vbCr
	End If
	
	If IsArray(E6_b_biz_partner) Then
		Response.Write "	.frm1.txtBpCd.Value			= """ & ConvSPChars(E6_b_biz_partner(A291_E6_bp_cd))	& """" & vbCr
		Response.Write "	.frm1.txtBpNm.Value			= """ & ConvSPChars(E6_b_biz_partner(A291_E6_bp_nm))	& """" & vbCr
	End If
	
    Response.Write "	.frm1.txtBankAcct.Value			= """ & ConvSPChars(E10_b_bank_acct_bank_acct_no)		& """" & vbCr    
    Response.Write "	.frm1.txtInputTypeNM.Value		= """ & ConvSPChars(E1_b_minor_minor_nm)				& """" & vbCr
    Response.Write "	.frm1.txtGlNo.value				= """ & ConvSPChars(E7_a_gl_gl_no)						& """" & vbCr
    
	Response.Write "	.ggoSpread.Source = .frm1.vspdData			" & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & strData	   & """" & vbCr
	
	Response.Write "	.ggoSpread.Source = .frm1.vspdData1				" & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & strData1	   & """" & vbCr
        
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
