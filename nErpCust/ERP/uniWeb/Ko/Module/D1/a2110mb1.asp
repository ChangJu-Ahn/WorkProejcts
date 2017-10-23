<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<% 
	Call LoadBasisGlobalInf()
	Dim lgOpModeCRUD

	On Error Resume Next
	Err.Clear

    Call HideStatusWnd

    lgOpModeCRUD      = Request("txtMode")
    
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
	On Error Resume Next
	Err.Clear

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
	On Error Resume Next
	Err.Clear
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
	On Error Resume Next
	Err.Clear
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iPD1G050
    
    Dim iStrNextKey
    Dim iStrData
    Dim iStrMo_cd
    Dim export_group
    Dim import_group
    Dim export_nm
    Dim iLngRow
    Dim iIntLoopCount 
    Dim iIntQueryCount
    
    Const C_SHEETMAXROWS_D  = 100
    Const C_QueryConut = 0
    Const C_MaxFetchRc = 1
    Const C_import_key = 2

    Const A542_EG1_E1_a_acct_trans_type_mo_cd = 0
    Const A542_EG1_E1_a_acct_trans_type_mo_nm = 1
    Const A542_EG1_E2_b_acct_conf_batch_fg = 2
    Const A542_EG1_E2_b_acct_conf_batch_fg_nm = 3
    Const A542_EG1_E2_b_acct_conf_gl_posting_fg = 4
    Const A542_EG1_E2_b_acct_conf_gl_posting_fg_nm = 5
    Const A542_EG1_E2_b_company_inv_post_fg = 6

    ReDim import_group(2)

    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear																		'☜: Clear Error status
    
    If Trim(Request("iPageNo")) = "" Then 
        iIntQueryCount = 0       
    Else    
        iIntQueryCount = Request("iPageNo")											'쿼리카운트 받아오고 
    End If
    
    If Trim(Request("iStrNextKey")) = "" Then										'넥스트키값이 없으면 
		iStrMo_Cd = Request("txtMO_Cd")												'입력받은 코드값을 받고   
    Else																			'넥스트키값이 있으면 
		iStrMo_cd = iStrNextKey														'넥스트 키값을 세팅 
    End If

    Set iPD1G050 = Server.CreateObject("PD1G050.cBListAcctCnfmSvr")

    If CheckSYSTEMError(Err, True) = True Then     
		Exit Sub
    End If    

    import_group(C_QueryConut) = iIntQueryCount
    import_group(C_MaxFetchRc) = C_SHEETMAXROWS_D
    import_group(C_import_key) = iStrMo_cd

    Call iPD1G050.B_List_Acct_Conf_Svr(gStrGlobalCollection, import_group, export_group, export_nm)

    If CheckSYSTEMError(Err, True) = True Then     
		Set iPD1G050 = Nothing
		Exit Sub
    End If    

    Set iPD1G050 = Nothing

    iStrData = "" 
    iIntLoopCount = 0
	iStrNextKey = ""

    For iLngRow = 0 To UBound(export_group, 1)
        iIntLoopCount = iIntLoopCount + 1 
		If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, A542_EG1_E1_a_acct_trans_type_mo_cd)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, A542_EG1_E1_a_acct_trans_type_mo_nm)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, A542_EG1_E2_b_acct_conf_batch_fg)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, A542_EG1_E2_b_acct_conf_batch_fg_nm)))			
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, A542_EG1_E2_b_acct_conf_gl_posting_fg)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, A542_EG1_E2_b_acct_conf_gl_posting_fg_nm)))			
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(export_group(iLngRow, A542_EG1_E2_b_company_inv_post_fg)))
			iStrData = iStrData & Chr(11) & Cstr(Cint(iIntQueryCount)*Cint(C_SHEETMAXROWS_D) + iLngRow + 1)
			iStrData = iStrData & Chr(11) & Chr(12)         			
		Else
			iStrNextKey = export_group(UBound(export_group, 1), A542_EG1_E1_a_acct_trans_type_mo_cd)
			iIntQueryCount = iIntQueryCount + 1
		End If
    Next

    Response.Write " <Script Language=vbscript>									 " & vbCr
    Response.Write " With parent											     " & vbCr
    Response.Write " .ggoSpread.Source    = .frm1.vspdData					     " & vbCr     
    Response.Write " .ggoSpread.SSShowData  """ & iStrData					& """" & vbCr
    Response.Write " .igPageNo            = """ & iIntQueryCount			& """" & vbCr
    Response.Write " .igStrNextKey        = """ & ConvSPChars(iStrNextKey)  & """" & vbCr
    Response.Write " .frm1.txtMo_Nm.value = """ & ConvSPChars(export_nm)    & """" & vbCr
    Response.Write " .DbQueryOk													 " & vbCr
    Response.Write " End With													 " & vbCr
    Response.Write "</Script>													 " & vbCr
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    Dim iPD1G050
    Dim iErrPosition
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Set iPD1G050 = Server.CreateObject("PD1G050.cBMngAcctCnfmSvr")

    If CheckSYSTEMError(Err, True) = True Then     
       Set iPD1G050 = Nothing
       Exit Sub
    End If    
 
    Call iPD1G050.B_MANAGE_ACCT_CONF_SVR(gStrGlobalCollection, Trim(Request("txtSpread")),iErrPosition)  

    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then     
       Set iPD1G050 = Nothing
       Exit Sub
    End If    
    
    Set iPD1G050 = Nothing
 
    Response.Write " <Script Language=vbscript> " & vbCr
    Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
End Sub    

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
	On Error Resume Next
	Err.Clear
End Sub

%>
