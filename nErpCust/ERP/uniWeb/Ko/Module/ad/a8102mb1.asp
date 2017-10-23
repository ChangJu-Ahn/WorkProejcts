<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%   
    Dim lgOpModeCRUD

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	Call LoadBasisGlobalInf()
    Call loadInfTB19029B("I", "A", "NOCOOKIE", "MB")
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
    lgOpModeCRUD      = Request("txtMode")                                          '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD    
        Case CStr(UID_M0001)                                                         '☜: Query
          '  Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
          '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
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
    
    '-------------------------------------------------------------------------------------------------
	'IMPORTS View 상수 
	'-------------------------------------------------------------------------------------------------
    Const A390_I3_org_change_id		= 0							'View Name : in_dept_cd b_acct_dept
    Const A390_I3_dept_cd			= 1

    Const A390_I4_temp_gl_no		= 0							'View Name : in_temp_gl a_temp_gl
    Const A390_I4_temp_gl_dt		= 1
    Const A390_I4_gl_type			= 2
    Const A390_I4_insrt_user_id		= 3
    Const A390_I4_updt_user_id		= 4
    Const A390_I4_issued_dt			= 5
    Const A390_I4_gl_input_type		= 6
    Const A390_I4_cr_amt			= 7
    Const A390_I4_cr_loc_amt		= 8
    Const A390_I4_dr_amt			= 9
    Const A390_I4_dr_loc_amt		= 10
    Const A390_I4_conf_fg			= 11
    Const A390_I4_insrt_dt			= 12
    Const A390_I4_updt_dt			= 13
    Const A390_I4_hq_brch_fg		= 14
    Const A390_I4_internal_cd		= 15

	Dim obj3PADG010
	
	Dim iCommandSent
	Dim I3_b_acct_dept
	Dim I4_a_temp_gl
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear          
	
	Redim I3_b_acct_dept(A390_I3_dept_cd)
	Redim I4_a_temp_gl(A390_I4_internal_cd)
	
	iCommandSent							= "DELETE"
	I3_b_acct_dept(A390_I3_org_change_id)	= Request("txtOrgChangeId")
	I3_b_acct_dept(A390_I3_dept_cd)			= Request("txtDeptCd")
	I4_a_temp_gl(A390_I4_temp_gl_no)		= Request("txtTempGlNo")


	Set obj3PADG010 = CreateObject("PADG010.cAMngTmpGlHqSvr")
    
    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If

	Call obj3PADG010.A_MANAGE_TEMP_GL_HQ_SVR (gStrGloBalCollection, iCommandSent, , , I3_b_acct_dept, I4_a_temp_gl)

	If CheckSYSTEMError(Err, True) = True Then
       Set obj3PADG010 = Nothing
       Exit Sub
    End If

	Set obj3PADG010 = Nothing

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
	'IMPORTS View 상수 
	'-------------------------------------------------------------------------------------------------
	Const A278_I7_biz_area_cd		= 0						'View Name : import_next_temp_gl_no a_temp_gl   'Not CoolGen
	Const A278_I7_temp_gl_no		= 1

	'-------------------------------------------------------------------------------------------------
	'EXPORTS View 상수 
	'-------------------------------------------------------------------------------------------------
	  'Group Name : export_grp_temp_cr
	Const A278_EG1_E1_biz_area_cd	= 0						'View Name : exp_grp_bizarea_cr b_biz_area
	Const A278_EG1_E1_biz_area_nm	= 1
	Const A278_EG1_E2_temp_gl_no	= 2						'View Name : exp_grp_tempgl_cr a_temp_gl
	Const A278_EG1_E2_temp_gl_dt	= 3
	Const A278_EG1_E2_ref_no		= 4
	Const A278_EG1_E2_dr_loc_amt	= 5
	Const A278_EG1_E2_hq_brch_no	= 6

	'Group Name : export_grp_temp_dr
	Const A278_EG2_E1_biz_area_cd	= 0						'View Name : exp_grp_bizarea_dr b_biz_area
	Const A278_EG2_E1_biz_area_nm	= 1
	Const A278_EG2_E2_temp_gl_no	= 2						'View Name : exp_grp_tempgl_dr a_temp_gl
	Const A278_EG2_E2_temp_gl_dt	= 3
	Const A278_EG2_E2_ref_no		= 4
	Const A278_EG2_E2_dr_loc_amt	= 5
	Const A278_EG2_E2_hq_brch_no	= 6

	Const A278_E1_temp_gl_no		= 0						'View Name : export_next_key a_temp_gl
	Const A278_E1_temp_gl_dt		= 1

	Const C_SHEETMAXROWS_D			= 100

	Dim obj1PADG010
	Dim strtxttab
	Dim LngRow
	Dim LngRow1
	Dim LngMaxRow
	Dim LngMaxRow1
	Dim LngMaxRow2
	Dim strtempglhqnodr
	Dim strtempglhqnocr
	Dim iStrData
	Dim iStrData1

	Dim I1_ief_supplied_select_char
	Dim I2_b_biz_area_biz_area_cd
	Dim I3_a_temp_gl_temp_gl_dt
	Dim I4_a_temp_gl_temp_gl_dt
	Dim I5_a_assign_acct_wk_loc_amt
	Dim I6_a_assign_acct_wk_loc_amt
	Dim I7_a_temp_gl
	Dim EG1_export_grp_temp_cr
	Dim EG2_export_grp_temp_dr
	Dim E1_a_temp_gl

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear          

	strtxttab	= Request("txttab")
	LngMaxRow	= UNIConvNum(Request("txtMaxRows1"),0)
	LngMaxRow1	= UNIConvNum(Request("txtMaxRows1"),0)
	LngMaxRow2	= UNIConvNum(Request("txtMaxRows2"),0)

	I1_ief_supplied_select_char  = Request("txttab")
	I2_b_biz_area_biz_area_cd    = Request("txtBizArea")
    I3_a_temp_gl_temp_gl_dt      = UNIConvDate(Request("txtFromdt"))      
    I4_a_temp_gl_temp_gl_dt      = UNIConvDate(Request("txtTodt"))
    I5_a_assign_acct_wk_loc_amt  = UNIConvNum(Request("txtfromamt"),0)
    I6_a_assign_acct_wk_loc_amt  = UNIConvNum(Request("txttoamt"),0)

	Set obj1PADG010 = CreateObject("PADG010.cAListTmpGlHqSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If

	Call obj1PADG010.A_LIST_TEMP_GL_HQ_SVR (gStrGloBalCollection, , I1_ief_supplied_select_char, I2_b_biz_area_biz_area_cd, _
											I3_a_temp_gl_temp_gl_dt, I4_a_temp_gl_temp_gl_dt, I5_a_assign_acct_wk_loc_amt, I6_a_assign_acct_wk_loc_amt, _
											I7_a_temp_gl, EG1_export_grp_temp_cr, EG2_export_grp_temp_dr, E1_a_temp_gl)

	
	If CheckSYSTEMError(Err, True) = True Then
       Set obj1PADG010 = Nothing
       Exit Sub
    End If    

	Set obj1PADG010 = Nothing

	If isArray(EG2_export_grp_temp_dr) then

		For LngRow = 0 To Ubound(EG2_export_grp_temp_dr)

			If LngRow = 0 Then        
				strtempglhqnodr = EG2_export_grp_temp_dr(LngRow, A278_EG2_E2_hq_brch_no)
			End If

			iStrData = iStrData & Chr(11) & ""																					'1 C_wkchk
			iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG2_export_grp_temp_dr(LngRow, A278_EG2_E2_temp_gl_dt))         '2 c_tempgldt
			iStrData = iStrData & Chr(11) & ConvSPChars(EG2_export_grp_temp_dr(LngRow, A278_EG2_E1_biz_area_cd))				'3 C_bizareacd
			iStrData = iStrData & Chr(11) & ConvSPChars(EG2_export_grp_temp_dr(LngRow, A278_EG2_E1_biz_area_nm))				'4 C_bizareanm
			iStrData = iStrData & Chr(11) & ConvSPChars(EG2_export_grp_temp_dr(LngRow, A278_EG2_E2_temp_gl_no))  				'5 C_tempglno
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG2_export_grp_temp_dr(LngRow, A278_EG2_E2_dr_loc_amt), ggAmtOfMoney.DecPoint, 0)     	'6 C_drlocamt
	        iStrData = iStrData & Chr(11) & LngMaxRow + LngRow + 1
	        iStrData = iStrData & Chr(11) & Chr(12)

	    Next

	End If

	If isArray(EG1_export_grp_temp_cr) Then

		For LngRow1 = 0 To Ubound(EG1_export_grp_temp_cr)

			strtempglhqnocr = EG1_export_grp_temp_cr(LngRow1, A278_EG1_E2_hq_brch_no)

			If Trim(strtxttab) = "2" Then
				If  Trim(strtempglhqnodr) = Trim(strtempglhqnocr) Then 
					iStrData1 = iStrData1 & Chr(11) & ""		'1 C_wkchk
					iStrData1 = iStrData1 & Chr(11) & UNIDateClientFormat(EG1_export_grp_temp_cr(LngRow1, A278_EG1_E2_temp_gl_dt))	'2 c_tempgldt
					iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_grp_temp_cr(LngRow1, A278_EG1_E1_biz_area_cd))			'3 C_bizareacd
					iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_grp_temp_cr(LngRow1, A278_EG1_E1_biz_area_nm))			'4 C_bizareanm
					iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_grp_temp_cr(LngRow1, A278_EG1_E2_temp_gl_no))			'5 C_tempglno
					iStrData1 = iStrData1 & Chr(11) & UNINumClientFormat(EG1_export_grp_temp_cr(LngRow1, A278_EG1_E2_dr_loc_amt), ggAmtOfMoney.DecPoint, 0)    	'6 C_drlocamt
					iStrData1 = iStrData1 & Chr(11) & LngMaxRow + LngRow1 + 1
					iStrData1 = iStrData1 & Chr(11) & Chr(12)
				End If
			Else
				iStrData1 = iStrData1 & Chr(11) & ""		'1 C_wkchk
				iStrData1 = iStrData1 & Chr(11) & UNIDateClientFormat(EG1_export_grp_temp_cr(LngRow1, A278_EG1_E2_temp_gl_dt))	'2 c_tempgldt
				iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_grp_temp_cr(LngRow1, A278_EG1_E1_biz_area_cd))			'3 C_bizareacd
				iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_grp_temp_cr(LngRow1, A278_EG1_E1_biz_area_nm))			'4 C_bizareanm
				iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_grp_temp_cr(LngRow1, A278_EG1_E2_temp_gl_no))  	    '5 C_tempglno
				iStrData1 = iStrData1 & Chr(11) & UNINumClientFormat(EG1_export_grp_temp_cr(LngRow1, A278_EG1_E2_dr_loc_amt), ggAmtOfMoney.DecPoint, 0)     	'6 C_drlocamt
				iStrData1 = iStrData1 & Chr(11) & LngMaxRow + LngRow1 + 1
				iStrData1 = iStrData1 & Chr(11) & Chr(12)
			End If

		Next

	End If

	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write " With parent				" & vbCr

	If IsArray(EG2_export_grp_temp_dr) then
		If Trim(Request("txtBizArea")) <> "" Then
			 Response.Write "	.frm1.txtBizArea.Value   = 	""" & ConvSPChars(EG2_export_grp_temp_dr(0,A278_EG2_E1_biz_area_cd))		& """" & vbCr
			 Response.Write "	.frm1.txtBizAreaNm.Value = 	""" & ConvSPChars(EG2_export_grp_temp_dr(0,A278_EG2_E1_biz_area_nm))		& """" & vbCr
		End If
	End If
	
	If Trim(strtxttab) = "1" Then
			Response.Write " 	.ggoSpread.Source = .frm1.vspdData			"	& vbCr
	Else
		If Trim(strtxttab) = "2" Then
			Response.Write " 	.ggoSpread.Source = .frm1.vspdData6			"	& vbCr
		End If
	End If
			Response.Write " 	.ggoSpread.SSShowData	""" & iStrData	& """"	& vbCr

	If Trim(strtxttab) = "1" Then
			Response.Write " 	.ggoSpread.Source = .frm1.vspdData2			"	& vbCr
	Else
		If Trim(strtxttab) = "2" Then
			Response.Write " 	.ggoSpread.Source = .frm1.vspdData7			"	& vbCr
		End If
	End If
			Response.Write " 	.ggoSpread.SSShowData	""" & iStrData1		& """"	& vbCr
			Response.Write " 	.DbQueryOk				""" & strtxttab		& """"	& vbCr

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
    'Group Name : import_grp_temp_gl
    Const A054_IG1_I1_temp_gl_no = 0						'View Name : import_grp_temp_gl_no a_temp_gl
	
	Dim obj2PADG010
	Dim LngMaxRow
	Dim LngRow
	Dim strtxttab
	Dim arrTemp
	Dim arrVal

	Dim I1_ief_supplied_select_char
	Dim IG1_import_grp_temp_gl
	Dim I2_a_gl_updt_user_id

	On Error Resume Next
	Err.Clear

    LngMaxRow						= UNIConvNum(Request("txtMaxRows"),0)
    strtxttab						= Request("txtstrtab")

    Redim IG1_import_grp_temp_gl(LngMaxRow - 1, A054_IG1_I1_temp_gl_no)
    I1_ief_supplied_select_char		= Request("txtstrtab")
    I2_a_gl_updt_user_id			= Request("txtUpdtUserId")

    arrTemp							= Split(Request("txtSpread"), gRowSep)

	For LngRow = 0 To LngMaxRow - 1
        arrVal = Split(arrTemp(LngRow), gColSep)								'☜: Group Count
    	IG1_import_grp_temp_gl(LngRow,A054_IG1_I1_temp_gl_no) = arrVal(0)		' ItemSEQ  * Key
	Next

	Set obj2PADG010 = CreateObject("PADG010.cAConnectTmpGlSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If

	Call obj2PADG010.A_CONNECT_TEMP_GL_SVR (gStrGloBalCollection, I1_ief_supplied_select_char, IG1_import_grp_temp_gl, _
											I2_a_gl_updt_user_id)

	If CheckSYSTEMError(Err, True) = True Then
       Set obj2PADG010 = Nothing
       Exit Sub
    End If

	Set obj2PADG010 = Nothing

	Response.Write "<Script Language=vbscript>				" & vbcr
	Response.Write " With Parent							" & vbCr
	Response.Write " 	.DbSaveOk							" & vbCr
	Response.Write " End With								" & vbCr
	Response.Write "</Script>								" & vbCr

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
