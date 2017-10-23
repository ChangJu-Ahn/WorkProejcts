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
    Call LoadInfTB19029B("I", "A", "NOCOOKIE", "MB")
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
'	        Call SubBizDelete()
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
    'Group Name : export_grp_temp_cr
    Const A351_EG1_E1_biz_area_cd	= 0					'View Name : exp_grp_biz_area_cr b_biz_area
    Const A351_EG1_E1_biz_area_nm	= 1
    Const A351_EG1_E2_temp_gl_no	= 2					'View Name : exp_grp_tempgl_cr a_temp_gl
    Const A351_EG1_E2_temp_gl_dt	= 3
    Const A351_EG1_E2_ref_no		= 4
    Const A351_EG1_E2_dr_loc_amt	= 5
    Const A351_EG1_E2_hq_brch_no	= 6
    Const A351_EG1_E2_dr_amt		= 7
    Const A351_EG1_E2_temp_gl_desc	= 8					'View Name : exp_grp_tempgl_cr a_temp_gl	>>air

    'Group Name : export_grp_item
    Const A351_EG2_E1_item_seq		= 0					'View Name : exp_grp_item_a_temp_gl_item a_temp_gl_item
    Const A351_EG2_E1_dr_cr_fg		= 1
    Const A351_EG2_E1_item_loc_amt	= 2
    Const A351_EG2_E2_temp_gl_no	= 3					'View Name : exp_grp_item_a_temp_gl a_temp_gl
    Const A351_EG2_E3_acct_cd		= 4					'View Name : exp_grp_item_a_acct a_acct
    Const A351_EG2_E3_acct_nm		= 5
    Const A351_EG2_E4_major_cd		= 6					'View Name : exp_grp_item_b_major b_major
    Const A351_EG2_E5_minor_cd		= 7					'View Name : exp_grp_item_b_minor b_minor
    Const A351_EG2_E5_minor_nm		= 8
    Const A351_EG2_E1_item_desc 	= 9   				'20080403   결의전표항목(조회-적요)    >>air

	Dim obj1PADG010
	Dim strgubun
	Dim strtab
	Dim strtempglno
	Dim strqrychk
	Dim GroupCnt
	Dim GroupCnt1
	Dim LngRow
	Dim LngRow1
	Dim LngMaxRow
	Dim iStrData
	Dim iStrData1

	Dim I1_ief_supplied_select_char
	Dim I2_a_temp_gl_temp_gl_no
	Dim I3_a_temp_gl_item_item_seq
	Dim EG1_export_grp_temp_cr
	Dim EG2_export_grp_item
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear          

	strgubun	= Request("txtgubun")
	strtab		= Request("txtstrtab")
	strtempglno	= Request("txtTempGlNo")
	strqrychk	= request("txtquerychk") 
	LngMaxRow	= UNIConvNum(Request("txtMaxRows"),0)

	I2_a_temp_gl_temp_gl_no = strtempglno

    If Trim(strtab) = "2"   Then
        If Trim(strqrychk)   = "Y" Then
           I1_ief_supplied_select_char =  strgubun
        Else
           I1_ief_supplied_select_char =  "1"
        End If

    Else
       I1_ief_supplied_select_char =  strtab
    End If

	Set obj1PADG010 = CreateObject("PADG010_TEST.cALkUpTmpGlItmSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If

	Call obj1PADG010.A_LOOKUP_TEMP_GL_ITEM_SVR (gStrGloBalCollection, I1_ief_supplied_select_char, I2_a_temp_gl_temp_gl_no, _
												I3_a_temp_gl_item_item_seq, EG1_export_grp_temp_cr, EG2_export_grp_item)

	If CheckSYSTEMError(Err, True) = True Then
       Set obj1PADG010 = Nothing
       Exit Sub
    End If

	Set obj1PADG010 = Nothing

	GroupCnt  = 0
	GroupCnt1 = 0

	If IsArray(EG2_export_grp_item) Then

		GroupCnt = Ubound(EG2_export_grp_item)

		For LngRow = 0 To Ubound(EG2_export_grp_item)

			iStrData = iStrData & Chr(11) & ConvSPChars(EG2_export_grp_item(LngRow, A351_EG2_E1_item_seq))
			iStrData = iStrData & Chr(11) & ConvSPChars(EG2_export_grp_item(LngRow, A351_EG2_E3_acct_nm))
			iStrData = iStrData & Chr(11) & ConvSPChars(EG2_export_grp_item(LngRow, A351_EG2_E5_minor_nm))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG2_export_grp_item(LngRow, A351_EG2_E1_item_loc_amt), ggAmtOfMoney.DecPoint, 0)
			iStrData = iStrData & Chr(11) & ConvSPChars(EG2_export_grp_item(LngRow, A351_EG2_E1_item_desc))	'결의전표항목(조회-적요) >>air
			iStrData = iStrData & Chr(11) & LngMaxRow + LngRow + 1
			iStrData = iStrData & Chr(11) & Chr(12)

		Next

	End If

	If IsArray(EG1_export_grp_temp_cr) then

		GroupCnt1 = Ubound(EG1_export_grp_temp_cr)

		For LngRow1 = 0 To Ubound(EG1_export_grp_temp_cr)

			iStrData1 = iStrData1 & Chr(11) & ""
			iStrData1 = iStrData1 & Chr(11) & UNIDateClientFormat(EG1_export_grp_temp_cr(LngRow1, A351_EG1_E2_temp_gl_dt))
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_grp_temp_cr(LngRow1, A351_EG1_E1_biz_area_cd))
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_grp_temp_cr(LngRow1, A351_EG1_E1_biz_area_nm))
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_grp_temp_cr(LngRow1, A351_EG1_E2_temp_gl_no))
			iStrData1 = iStrData1 & Chr(11) & UNINumClientFormat(EG1_export_grp_temp_cr(LngRow1, A351_EG1_E2_dr_loc_amt), ggAmtOfMoney.DecPoint, 0)
			iStrData1 = iStrData1 & Chr(11) & UNINumClientFormat(EG1_export_grp_temp_cr(LngRow1, A351_EG1_E2_dr_amt), ggAmtOfMoney.DecPoint, 0)
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG1_export_grp_temp_cr(LngRow1, A351_EG1_E2_temp_gl_desc))	'>>air
			iStrData1 = iStrData1 & Chr(11) & LngMaxRow + LngRow1 + 1
			iStrData1 = iStrData1 & Chr(11) & Chr(12)

		Next

	End If

	Response.Write "<Script Language=vbscript>	" & vbcr

	Response.Write " With Parent				" & vbCr

	If Trim(strtab) = "1" Then

		If strgubun = "1" Then

			Response.Write " 	.ggoSpread.Source = .frm1.vspdData3					" & vbcr
			Response.Write " 	Call parent.ggoSpread.ClearSpreadData()				" & vbcr
			Response.Write " 	.ggoSpread.SSShowData	""" &	iStrData	& """	" & vbCr
			Response.Write " 	.DbQueryOk2				""" &	strtab		& """	" & vbCr

		Else
			Response.Write " 	.frm1.vspdData4.MaxRows = 0							" & vbcr
			Response.Write " 	.ggoSpread.Source = .frm1.vspdData4					" & vbcr
			Response.Write " 	.ggoSpread.SSShowData	""" &	iStrData	& """	" & vbCr
			Response.Write " 	.DbQueryOk3											" & vbcr
		End If

	Else

		If strgubun = "1" Then

			Response.Write " 	.ggoSpread.Source = .frm1.vspdData8					" & vbcr
			Response.Write " 	Call parent.ggoSpread.ClearSpreadData()				" & vbcr
			Response.Write " 	.ggoSpread.SSShowData	""" &	iStrData	& """	" & vbCr
			Response.Write " 	.DbQuery6	" & """" & strtab & """" & "," & """" & "2" & """" & "," & """" & strtempglno & """" & vbCr

		Else

			Response.Write " 	.ggoSpread.Source = .frm1.vspdData9					" & vbcr
			Response.Write " 	Call parent.ggoSpread.ClearSpreadData()				" & vbcr
			Response.Write " 	.ggoSpread.SSShowData """ &		iStrData	& """	" & vbCr
			Response.Write " 	.DbQueryOk3											" & vbcr
		End If

	End If

	If  GroupCnt1 >= 0 And IsArray(EG1_export_grp_temp_cr) Then

		Response.Write " 	.ggoSpread.Source = .frm1.vspdData7						" & vbcr
		Response.Write " 	Call parent.ggoSpread.ClearSpreadData()					" & vbcr
		Response.Write " 	.ggoSpread.SSShowData """ &		iStrData1		& """	" & vbCr

	End if
	Response.Write " 	.DbQueryOk3				" & vbCr
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
	On Error Resume Next
	Err.Clear

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
