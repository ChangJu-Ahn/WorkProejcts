<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
     
    On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear 

    Dim lgOpModeCRUD
	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 
    
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "A", "COOKIE", "MB")
    Call LoadBNumericFormatB("I", "A","COOKIE","MB") 
                                                                          '☜: Clear Error status
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
	' 권한관리 추가 
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))    
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
          '  Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
          '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             'Call SubBizDelete()
             Call SubBizDeleteMulti()
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
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next
	Err.Clear

	Const A347_I1_a_temp_gl_temp_gl_no = 0
	
	Const A352_E1_a_temp_gl_temp_gl_no = 0
	Const A352_E1_a_temp_gl_temp_gl_dt = 1
	Const A352_E1_a_temp_gl_gl_type = 2
	Const A352_E1_a_temp_gl_input_type = 3
	Const A352_E1_a_temp_gl_cr_amt = 4
	Const A352_E1_a_temp_gl_cr_loc_amt = 5
	Const A352_E1_a_temp_gl_dr_amt = 6
	Const A352_E1_a_temp_gl_dr_loc_amt = 7
	'Const A352_E1_a_temp_gl_conf_fg = 8
	Const A352_E1_a_temp_gl_temp_gl_desc = 8
	Const A352_E1_a_temp_gl_project_no = 9
	Const A352_E1_a_temp_gl_org_change_id = 10
	Const A352_E1_a_temp_gl_dept_cd = 11
	Const A352_E1_dept_nm = 12
	
	Const A352_EG1_a_temp_gl_item_item_seq		= 0
	Const A352_EG1_a_temp_gl_item_dept_cd		= 1
	Const A352_EG1_a_temp_gl_item_dept_nm		= 2
	Const A352_EG1_a_temp_gl_item_acct_cd		= 3
	Const A352_EG1_a_temp_gl_item_acct_nm		= 4
	Const A352_EG1_a_temp_gl_item_dr_cr_fg		= 5	
	Const A352_EG1_a_temp_gl_item_item_amt		= 6
	Const A352_EG1_a_temp_gl_item_item_loc_amt	= 7
	Const A352_EG1_a_temp_gl_item_vat_type		= 8
	Const A352_EG1_a_temp_gl_item_item_desc		= 9
	Const A352_EG1_a_temp_gl_item_xch_rate		= 10
	Const A352_EG1_a_temp_gl_item_doc_cur		= 11
	Const A352_EG1_a_temp_gl_item_project_no		= 12
	Const A352_EG1_a_temp_gl_item_gl_no			= 13
	Const A352_EG1_a_temp_gl_item_org_change_id	= 14
	Const A352_EG1_a_temp_gl_item_acct_type		= 15
	Const A347_EG1_biz_area_cd = 16
	Const A347_EG1_biz_area_nm = 17
	Const A347_EG1_mgnt_fg = 18
'	Const A347_EG1_bal_fg = 19
	
	Dim PAUG020_cALkUpClsGlSvr
	Dim iStrData
	Dim iStrData1
    Dim iexportData
    Dim iexportData1
    Dim iLngRow
    Dim iLngCol
    Dim iStrCurrency
    
    Dim iStrTempGlNo

    Redim iStrTempGlNo(A347_I1_a_temp_gl_temp_gl_no+4)
    iStrTempGlNo(A347_I1_a_temp_gl_temp_gl_no)   = UCase(Trim(Request("txtTempGlNo")))
	iStrTempGlNo(A347_I1_a_temp_gl_temp_gl_no+1) = lgAuthBizAreaCd
	iStrTempGlNo(A347_I1_a_temp_gl_temp_gl_no+2) = lgInternalCd
	iStrTempGlNo(A347_I1_a_temp_gl_temp_gl_no+3) = lgSubInternalCd
	iStrTempGlNo(A347_I1_a_temp_gl_temp_gl_no+4) = lgAuthUsrID    
    
    Set PAUG020_cALkUpClsGlSvr = Server.CreateObject("PAUG020.cALkUpClsGlSvr")
    
    If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If
    
	Call PAUG020_cALkUpClsGlSvr.A_LOOKUP_CLS_GL_SVR(gStrGloBalCollection, iStrTempGlNo, iexportData, iexportData1)
	
	If CheckSYSTEMError(Err, True) = True Then					
		Set PAUG020_cALkUpClsGlSvr = Nothing        
		Exit Sub
    End If    

    Set PAUG020_cALkUpClsGlSvr = Nothing
    
	iStrCurrency = ConvSPChars(iexportData1(0, A352_EG1_a_temp_gl_item_doc_cur))
    iStrData	= ""	
    iStrData1	= ""

	For iLngRow = 0 To UBound(iexportData1, 1) 
		iStrData = iStrData & Chr(11) & iexportData1(iLngRow, A352_EG1_a_temp_gl_item_item_seq)
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_dept_cd))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_dept_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_acct_cd))
		iStrData = iStrData & Chr(11) & ""		
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_acct_nm))
		iStrData = iStrData & Chr(11) & UCase(ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_dr_cr_fg)))
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & iexportData1(iLngRow, A352_EG1_a_temp_gl_item_doc_cur)
		iStrData = iStrData & Chr(11) & UNINumClientFormat(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_xch_rate),ggExchRate.DecPoint, 0)
		iStrData = iStrData & Chr(11) & ""'XX
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_item_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_item_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_item_desc))
		iStrData = iStrData & Chr(11) & ""	
		iStrData = iStrData & Chr(11) & ""
		iStrData = iStrData & Chr(11) & ConvSPChars(iexportData1(iLngRow, A352_EG1_a_temp_gl_item_acct_cd))		
		iStrData = iStrData & Chr(11) & iLngRow + 1
		
		iStrData = iStrData & Chr(11) & Chr(12)		
		
		iStrData1 = iStrData1 & iStrData
		iStrData = ""
	Next
	
	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " With Parent " & vbCr		
	Response.Write " 	.frm1.txtTempGlNo.value		= """ & ConvSPChars(UCase(Trim(iexportData(A352_E1_a_temp_gl_temp_gl_no))))														  & """" & vbCr
	Response.Write "	.frm1.txtTempGLDt.Text		= """ & UNIDateClientFormat(iexportData(A352_E1_a_temp_gl_temp_gl_dt))															  & """" & vbCr
	Response.Write " 	.frm1.cboGLtype.value		= """ & iexportData(A352_E1_a_temp_gl_gl_type)																					  & """" & vbCr
	Response.Write " 	.frm1.cboGlInputType.Value	= """ & ConvSPChars(iexportData(A352_E1_a_temp_gl_input_type))																	  & """" & vbCr
	Response.Write " 	.frm1.txtCrLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A352_E1_a_temp_gl_cr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
	Response.Write " 	.frm1.txtDrLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iexportData(A352_E1_a_temp_gl_dr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr							
	Response.Write " 	.frm1.txtDesc.Value			= """ & ConvSPChars(iexportData(A352_E1_a_temp_gl_temp_gl_desc))				                                                  & """" & vbCr
	Response.Write " 	.frm1.txtDeptCd.value		= """ & UCase(Trim(iexportData(A352_E1_a_temp_gl_dept_cd)))				                                                          & """" & vbCr
	Response.Write " 	.frm1.txtDeptNm.value		= """ & ConvSPChars(iexportData(A352_E1_dept_nm))				                                                                  & """" & vbCr	
	Response.Write " 	.frm1.txtDocCur.value		= """ & iStrCurrency                                                                                                              & """" & vbCr
	Response.Write " 	.frm1.htxtTempGlNo.value	= """ & ConvSPChars(UCase(Trim(iexportData(A352_E1_a_temp_gl_temp_gl_no))))			                                              & """" & vbCr
	Response.Write " 	.ggoSpread.Source = .frm1.vspdData	      " & vbCr
	Response.Write " 	.ggoSpread.SSShowData """ & iStrData1 & """,""F""" & vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & -1 & "," & -1 & " ,.C_DocCur ,.C_ExchRate ,""D"" ,""I"",""X"",""X"")"	& vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & -1 & "," & -1 & " ,.C_DocCur ,.C_BalAmt   ,""A"" ,""I"",""X"",""X"")"	& vbCr
	Response.Write " Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & -1 & "," & -1 & " ,.C_DocCur ,.C_ItemAmt  ,""A"" ,""I"",""X"",""X"")"	& vbCr				
	Response.Write " 	.DbQueryOk								  " & vbCr
	Response.Write " End With " & vbCr
	Response.Write "</Script> " & vbCr 
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next	
	Err.Clear
		
	Const A393_I2_a_temp_gl_temp_gl_no = 0
    Const A393_I2_a_temp_gl_temp_gl_dt = 1
    Const A393_I2_a_temp_gl_org_change_id = 2
    Const A393_I2_a_temp_gl_dept_cd = 3
    Const A393_I2_a_temp_gl_gl_type = 4
    Const A393_I2_a_temp_gl_gl_input_type = 5
    Const A393_I2_a_temp_gl_temp_gl_desc = 6
    Const A393_I2_a_temp_gl_project_no = 7

	Dim PAUG020_cAMngClsGlSvr
	Dim iCommandSent
	Dim I1_b_currency
	Dim I2_a_gl
	Dim txtSpread
	Dim txtSpread3
	Dim iStrRetTempGlNo
	
	Dim iLngMaxRow
	Dim iLngMaxRow3
	Dim iLngRow
	Dim iArrTemp1
	Dim iArrTemp2

	'--------------------------------------------------------------------
	'A_GL에 대한 정보  Setting
	'--------------------------------------------------------------------
	iCommandSent = Request("txtCommandMode")	'Spread Sheet 내용을 담고 있는 Element명 
	I1_b_currency = gCurrency

    ReDim I2_a_gl(7)
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_no)		= UCase(Trim(Request("txtTempGlNo")))	
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_dt)		= UNIConvDate(Request("txtTempGlDt"))
	I2_a_gl(A393_I2_a_temp_gl_org_change_id)	= Trim(Request("hOrgChangeId"))
	I2_a_gl(A393_I2_a_temp_gl_dept_cd)			= UCase(Trim(Request("txtDeptCd")))
	I2_a_gl(A393_I2_a_temp_gl_gl_type)			= Trim(Request("cboGlType")) 
	I2_a_gl(A393_I2_a_temp_gl_gl_input_type)	= Trim(Request("cboGlInputType"))      
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_desc)		= Request("txtDesc")

    Dim I3_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
    Const A393_I3_a_data_auth_data_BizAreaCd = 0
    Const A393_I3_a_data_auth_data_internal_cd = 1
    Const A393_I3_a_data_auth_data_sub_internal_cd = 2
    Const A393_I3_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I3_a_data_auth(3)
	I3_a_data_auth(A393_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I3_a_data_auth(A393_I3_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I3_a_data_auth(A393_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I3_a_data_auth(A393_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	'--------------------------------------------------------------------
	'A_GL_ITEM에 대한 정보  Setting
	'--------------------------------------------------------------------
	txtSpread = Request("txtSpread")
    
    '--------------------------------------------------------------------
	'A_GL_DTL에 대한 정보  Setting
	'--------------------------------------------------------------------
	txtSpread3 = Request("txtSpread3")
	
	'--------------------------------------------------------------------
	'실행하기.
	'--------------------------------------------------------------------	
		
	Set PAUG020_cAMngClsGlSvr = Server.CreateObject("PAUG020.cAMngClsGlSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If
	
	iStrRetTempGlNo = PAUG020_cAMngClsGlSvr.A_MANAGE_CLS_GL_SVR(gStrGloBalCollection, iCommandSent, I1_b_currency, I2_a_gl, txtSpread, txtSpread3,I3_a_data_auth) 	

	If CheckSYSTEMError(Err, True) = True Then		
		Set PAUG020_cAMngClsGlSvr = Nothing
		Exit Sub
    End If
    
    Set PAUG020_cAMngClsGlSvr  = Nothing
    
    If InStr(1,ConvSPChars(UCase(Trim(Request("txtTempGlNo")))),"'") > 0 then
		iStrRetTempGlNo = ConvSPChars(UCase(Trim(Request("txtTempGlNo"))))
	End If    

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.DbSaveOk """ & iStrRetTempGlNo	&			 """" & vbCr    
    Response.Write " End With											" & vbCr
    Response.Write " </Script>											" & vbCr
End Sub    

'============================================================================================================
' Name : SubBizDeleteMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizDeleteMulti()

	On Error Resume Next
	Err.Clear

	Const A393_I2_a_temp_gl_temp_gl_no = 0
    Const A393_I2_a_temp_gl_temp_gl_dt = 1
    Const A393_I2_a_temp_gl_org_change_id = 2
    Const A393_I2_a_temp_gl_dept_cd = 3
    Const A393_I2_a_temp_gl_gl_type = 4
    Const A393_I2_a_temp_gl_gl_input_type = 5
    Const A393_I2_a_temp_gl_temp_gl_desc = 6
    Const A393_I2_a_temp_gl_project_no = 7

	Dim PAUG020_cAMngClsGlSvr
	Dim iCommandSent
	Dim I1_b_currency
	Dim I2_a_gl
	Dim txtSpread
	Dim txtSpread3
	Dim iStrRetTempGlNo
	
	iCommandSent = "DELETE"

	ReDim I2_a_gl(7)
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_no)			= UCase(Trim(Request("txtTempGlNo")))
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_dt)			= UNIConvDate(Request("txtTempGlDt"))
	I2_a_gl(A393_I2_a_temp_gl_org_change_id)		= Trim(Request("txtOrgChangeId"))
	I2_a_gl(A393_I2_a_temp_gl_dept_cd)				= UCase(Trim(Request("txtDeptCd")))
	I2_a_gl(A393_I2_a_temp_gl_gl_type)				= Trim(Request("cboGlType")) 
	I2_a_gl(A393_I2_a_temp_gl_gl_input_type)		= Trim(Request("txtGlinputType"))      
	I2_a_gl(A393_I2_a_temp_gl_temp_gl_desc)			= Request("txtDesc") 

    Dim I3_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
    Const A393_I3_a_data_auth_data_BizAreaCd = 0
    Const A393_I3_a_data_auth_data_internal_cd = 1
    Const A393_I3_a_data_auth_data_sub_internal_cd = 2
    Const A393_I3_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I3_a_data_auth(3)
	I3_a_data_auth(A393_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I3_a_data_auth(A393_I3_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I3_a_data_auth(A393_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I3_a_data_auth(A393_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	Set PAUG020_cAMngClsGlSvr = Server.CreateObject("PAUG020.cAMngClsGlSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If
	
	iStrRetTempGlNo = PAUG020_cAMngClsGlSvr.A_MANAGE_CLS_GL_SVR(gStrGloBalCollection, iCommandSent, I1_b_currency, I2_a_gl, txtSpread, txtSpread3,I3_a_data_auth) 	
		
	If CheckSYSTEMError(Err, True) = True Then		
		Set PAUG020_cAMngClsGlSvr = Nothing
		Exit Sub
    End If
    
    Set PAUG020_cAMngClsGlSvr  = Nothing

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.DbDeleteOK										" & vbCr    
    Response.Write " End With											" & vbCr
    Response.Write " </Script>											" & vbCr
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
