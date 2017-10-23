<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4232mb1
'*  4. Program Name         : 유동성전환 
'*  5. Program Desc         : Register of Loan Change
'*  6. Comproxy List        : PAFG460
'*  7. Modified date(First) : 2002-12-27
'*  8. Modified date(Last)  : 2003-05-19
'*  9. Modifier (First)     : Ahn, do hyun
'* 10. Modifier (Last)      : Ahn, do hyun
'* 11. Comment              :
'=======================================================================================================

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd

On Error Resume Next
Err.Clear                                                               '☜: Protect system from crashing

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

Dim txtLoanNo
Dim lgOpModeCRUD,lgErrorStatus,lgErrorPos

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

	lgErrorStatus	= "NO"
	lgErrorPos		= ""                                                           '☜: Set to space
	lgOpModeCRUD	= Request("txtMode")					'☜ : 현재 상태를 받음 

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

Const C_SHEETMAXROWS_D   = 100

Select Case lgOpModeCRUD
     Case CStr(UID_M0001)                                                         '☜: Query
          Call SubBizQuery()

     Case CStr(UID_M0002)       
          Call SubBizSave()

     Case CStr(UID_M0003)                                                         '☜: Delete
          Call SubBizDelete()
End Select
   
Response.End 


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	Dim PAFG460LOOKUP	
	Dim I1_f_ln_info
	Dim E1_f_ln_info, EG1_export_group	
	Dim iLngRow,iLngCol
	Dim iIntLoopCount
	Dim iStrData
	Dim iStrPrevKey
	Dim iIntMaxRows
	Dim iIntQueryCount

    Const A851_f_ln_plan_fr_dt = 0
    Const A851_f_ln_plan_to_dt = 1
    Const A851_f_ln_plan_chg_dt = 2
    Const A851_f_ln_plan_chg_dt_to = 3
    Const A851_f_ln_plan_loan_no = 4
    Const A851_f_ln_plan_plan_acct_cd = 5
    Const A851_f_ln_plan_plan_acct_nm = 6
    Const A851_f_ln_plan_loan_fg = 7
    Const A851_f_ln_plan_biz_area_cd_from = 8
    Const A851_f_ln_plan_biz_area_cd_to = 9    

    Const EA_f_ln_info_loan_nm = 0
    Const EA_f_ln_info_acct_nm = 1
    Const EA_f_ln_info_biz_area_nm_from = 2           '///사업장관련 추가된 인터페이스 
    Const EA_f_ln_info_biz_area_nm_to = 3             '///사업장관련 추가된 인터페이스    
    
    Const EA_f_ln_plan_flt_conv_dt1 = 0
    Const EA_f_ln_plan_loan_no1 = 1
    Const EA_f_ln_plan_loan_nm1 = 2
    Const EA_f_ln_plan_plan_acct_cd1 = 3
    Const EA_f_ln_plan_plan_acct_nm1 = 4
    Const EA_f_ln_plan_loan_acct_cd1 = 5
    Const EA_f_ln_plan_loan_acct_nm1 = 6
    Const EA_f_ln_plan_pay_plan_dt1 = 7
    Const EA_f_ln_plan_doc_cur1 = 8
    Const EA_f_ln_plan_xch_rate1 = 9
    Const EA_f_ln_plan_plan_amt1 = 10
    Const EA_f_ln_plan_plan_loc_amt1 = 11
    Const EA_f_ln_plan_loan_dt1 = 12
    Const EA_f_ln_plan_due_dt1 = 13
    Const EA_f_ln_plan_loan_int_rate1 = 14
    Const EA_f_ln_plan_pay_obj1 = 15
    Const EA_f_ln_plan_temp_gl_no1 = 16
    Const EA_f_ln_plan_gl_no1 = 17
    Const EA_f_ln_plan_ref_no1 = 18
    Const EA_f_ln_plan_ref_seq1 = 19
    Const EA_f_ln_plan_rdp_cls_fg1 = 20
    Const EA_f_ln_plan_resl_fg1 = 21
    Const EA_f_ln_plan_conf_fg1 = 22

    iStrPrevKey		= Trim(Request("lgStrPrevKey"))        
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")

    I1_f_ln_info	= Split(Request("txtKeyStream"), gColSep)

'    If Len(Trim(iIntQueryCount))  Then                                        '☜ : Chnage Nextkey str into int value
'       If Isnumeric(iIntQueryCount) Then
'          iIntQueryCount = CInt(iIntQueryCount)
'       End If
'    Else   
'       iIntQueryCount = 0
'    End If

	'권한관리 추가 2006-08-11 JYK
	Redim Preserve I1_f_ln_info(A851_f_ln_plan_biz_area_cd_to+4)

	I1_f_ln_info(A851_f_ln_plan_biz_area_cd_to+1) = lgAuthBizAreaCd
	I1_f_ln_info(A851_f_ln_plan_biz_area_cd_to+2) = lgInternalCd
	I1_f_ln_info(A851_f_ln_plan_biz_area_cd_to+3) = lgSubInternalCd
	I1_f_ln_info(A851_f_ln_plan_biz_area_cd_to+4) = lgAuthUsrID

	Set PAFG460LOOKUP = Server.CreateObject("PAFG460.bFLkUpCvSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If

'	Response.Write I1_f_ln_info(A851_f_ln_plan_loan_fg)
'	Response.End 

	Call PAFG460LOOKUP.F_LOOKUP_CV_SVR(gStrGloBalCollection, iStrPrevKey, C_SHEETMAXROWS_D, I1_f_ln_info, E1_f_ln_info, EG1_export_group)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG460LOOKUP = nothing
		Response.Write " <Script Language=vbscript> "							& vbCr
		Response.Write " parent.frm1.txtBizAreaNm.value  = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_biz_area_nm_from)) & """" & vbCr
		Response.Write " parent.frm1.txtBizAreaNm1.value = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_biz_area_nm_to))   & """" & vbCr		
		Response.Write "	Parent.DbQueryOk "									& vbcr
		Response.Write " </Script> "											& vbCr
		Exit Sub
    End If
    Set PAFG460LOOKUP	 = nothing

	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write " parent.frm1.txtLOAN_NM.value    = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_loan_nm))          & """" & vbCr
	Response.Write " parent.frm1.txtBizAreaNm.value  = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_biz_area_nm_from)) & """" & vbCr
	Response.Write " parent.frm1.txtBizAreaNm1.value = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_biz_area_nm_to))   & """" & vbCr		
    Response.Write "</Script>               " & vbCr

	iStrData = ""

	If IsEmpty(EG1_export_group) = False Then
		For iLngRow = 0 To UBound(EG1_export_group, 1)
'			iIntLoopCount = iIntLoopCount + 1
'			If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
			iStrData = iStrData & Chr(11) & "0"
			If Trim(I1_f_ln_info(A851_f_ln_plan_chg_dt)) <> "" Then
				iStrData = iStrData & Chr(11) & UNIDateClientFormat(I1_f_ln_info(A851_f_ln_plan_chg_dt))
			Else
				iStrData = iStrData & Chr(11) & ""
			End If
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_plan_loan_no1)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_plan_loan_nm1)))
			iStrData = iStrData & Chr(11) & ConvSPChars(		I1_f_ln_info(A851_f_ln_plan_plan_acct_cd))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(		E1_f_ln_info(EA_f_ln_info_acct_nm))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_plan_loan_acct_cd1)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_plan_loan_acct_nm1)))
			iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow,EA_f_ln_plan_pay_plan_dt1))
			iStrData = iStrData & Chr(11) & ConvSPChars(		Trim(EG1_export_group(iLngRow,EA_f_ln_plan_doc_cur1)))
			iStrData = iStrData & Chr(11) & UNINumClientFormat( EG1_export_group(iLngRow,EA_f_ln_plan_xch_rate1)		,ggExchRate.DecPoint	,0)
			iStrData = iStrData & Chr(11) & UNINumClientFormat( EG1_export_group(iLngRow,EA_f_ln_plan_plan_amt1)		,ggAmtOfMoney.DecPoint	,0)
			iStrData = iStrData & Chr(11) & UNINumClientFormat( EG1_export_group(iLngRow,EA_f_ln_plan_plan_loc_amt1)		,ggAmtOfMoney.DecPoint	,0)
			iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow,EA_f_ln_plan_loan_dt1))
			iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow,EA_f_ln_plan_due_dt1))
			iStrData = iStrData & Chr(11) & UNINumClientFormat( EG1_export_group(iLngRow,EA_f_ln_plan_loan_int_rate1),ggExchRate.DecPoint	,0)
			iStrData = iStrData & Chr(11) & ConvSPChars(		Trim(EG1_export_group(iLngRow,EA_f_ln_plan_pay_obj1)))
		    iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
			iStrData = iStrData & Chr(11) & Chr(12)
'			Else
'				iStrPrevKey = EG1_export_group(UBound(EG1_export_group, 1), EA_f_ln_his_chg_dt2)
'				iIntQueryCount = iIntQueryCount + 1
'				Exit For
'			End If
		Next
		
'		If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
'			iStrPrevKey = ""
'		    iIntQueryCount = ""
'		End If

	End If

	Response.Write " <Script Language=vbscript> "							& vbCr
	Response.Write " With parent "											& vbCr
	Response.Write "	.ggoSpread.Source = .frm1.vspdData "				& vbcr
	Response.Write "    .frm1.vspdData.Redraw = False "						& vbCr
	Response.Write "	.ggoSpread.SSShowData  """ & iStrData & """ ,""F""" & vbCr
	Response.Write "	.lgPageNo			= """ & iIntQueryCount	& """"	& vbCr
	Response.Write "	.lgStrPrevKey		= """ & iStrPrevKey		& """"	& vbCr
	Response.Write "	Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & iIntMaxRows + 1 & "," & iIntMaxRows + iLngRow & ",.C_Doc_Cur, .C_PLAN_AMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
	Response.Write "	.DbQueryOk "										& vbcr
	Response.Write "    .frm1.vspdData.Redraw = True "						& vbCr
    Response.Write " End With "												& vbCr
    Response.Write " </Script> "											& vbCr

End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status
    
	Dim PAFG460CU
	Dim iarrData
	Dim I1_f_ln_info
	Dim txtSpread

    Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
    Const A659_I1_a_data_auth_data_BizAreaCd = 0
    Const A659_I1_a_data_auth_data_internal_cd = 1
    Const A659_I1_a_data_auth_data_sub_internal_cd = 2
    Const A659_I1_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A659_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(A659_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I1_a_data_auth(A659_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(A659_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	iarrData = Split(Request("txtSpread"), gRowSep)
    Set PAFG460CU = server.CreateObject ("PAFG460.cFMngCvSvr")   
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If

	Call PAFG460CU.F_MANAGE_CV_SVR(gStrGlobalCollection, iarrData,I1_a_data_auth)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG460CU = nothing
		Exit Sub	
    End If

    Set PAFG460CU = nothing

	Response.Write "<Script Language=vbscript>		" & vbCr
	Response.Write " parent.DbSaveOk()				" & vbCr
    Response.Write "</Script>						" & vbCr  
End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status    
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
