<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->

<%
     
    Dim lgOpModeCRUD
    Dim lgstrWkfg
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	Call LoadBasisGlobalInf()

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD		= Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgStrWkfg			= Request("htxtWorkFg")
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizSaveMulti()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizDeleteSaveMulti()
    End Select


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next
	Err.Clear	

		
	Const A891_I2_dept_cd			= 0
	Const A891_I2_org_change_id		= 1
	



	Dim PAGG116_cAMngOneBtchToGlSvr
	Dim pvCommandSent
	Dim I2_b_acct_dept
	Dim I3_gl_dt
	Dim I4_gl_input_type
	Dim IG1_a_batch_no
	
		
	pvCommandSent								= "SAVE"
	ReDim I2_b_acct_dept(A891_I2_org_change_id)
	I2_b_acct_dept(A891_I2_dept_cd)				= Trim(Request("txtDeptCd"))
	I2_b_acct_dept(A891_I2_org_change_id)		= Trim(Request("hOrgChangeId"))
	I3_gl_dt									= UNIConvDate(Request("GIDate"))
	I4_gl_input_type							= Trim(Request("txtTransType"))
	IG1_a_batch_no								= Trim(Request("txtSpread"))

	Set PAGG116_cAMngOneBtchToGlSvr = Server.CreateObject("PAGG116.cAMngOneBtchToGlSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If
	

	Call PAGG116_cAMngOneBtchToGlSvr.A_MANAGE_ONE_BATCH_TO_GL_SVR(gStrGloBalCollection, _
															pvCommandSent, _
															I2_b_acct_dept, _
															I3_gl_dt, _
															I4_gl_input_type, _
															IG1_a_batch_no) 
	
	
	If CheckSYSTEMError(Err, True) = True Then		
       Set PAGG116_cAMngOneBtchToGlSvr = Nothing
       Exit Sub
    End If
    
    Set PAGG116_cAMngOneBtchToGlSvr  = Nothing

	Response.Write " <Script Language=vbscript>			" & vbCr
	Response.Write " With parent						" & vbCr
	Response.Write "	.frm1.txtFromReqDt1.text = """ & .frm1.txtFromReqDt.text & """" & vbCr
	Response.Write "	.frm1.txtToReqDt1.text = """ & .frm1.txtToReqDt.text & """" & vbCr
	Response.Write "	.frm1.txtTransType1.value = """ & ConvSPChars(.frm1.txtTransType.value)	& """" & vbCr
	Response.Write "	.frm1.txtTransTypeNm1.value = """ & ConvSPChars(.frm1.txtTransTypeNm.value)	& """" & vbCr
	Response.Write "	.frm1.txtBizCd1.value = """ & ConvSPChars(.frm1.txtBizCd.value)	& """" & vbCr
	Response.Write "	.frm1.txtBizNm1.value = """ & ConvSPChars(.frm1.txtBizNm.value)	& """" & vbCr
	Response.Write "	.DbSaveOk   " & vbCr
	Response.Write " End With   " & vbCr
	Response.Write " </Script>  " & vbCr  
End Sub    


'============================================================================================================
' Name : SubBizDeleteSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizDeleteSaveMulti()


	On Error Resume Next
	Err.Clear	

		
	Const A891_I2_dept_cd			= 0
	Const A891_I2_org_change_id		= 1
	



	Dim PAGG116_cAMngOneBtchToGlSvr
	Dim pvCommandSent
	Dim I2_b_acct_dept
	Dim I3_gl_dt
	Dim I4_gl_input_type
	Dim IG1_a_batch_no
	
		
	pvCommandSent								= "DELETE"
	ReDim I2_b_acct_dept(A891_I2_org_change_id)
	I2_b_acct_dept(A891_I2_dept_cd)				= ""
	I2_b_acct_dept(A891_I2_org_change_id)		= ""
	I3_gl_dt									= "1900-01-01"
	I4_gl_input_type							= Trim(Request("txtTransType1"))
	IG1_a_batch_no								= Trim(Request("txtSpread"))

	Set PAGG116_cAMngOneBtchToGlSvr = Server.CreateObject("PAGG116.cAMngOneBtchToGlSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If
	
	Call PAGG116_cAMngOneBtchToGlSvr.A_MANAGE_ONE_BATCH_TO_GL_SVR(gStrGloBalCollection, _
															pvCommandSent, _
															I2_b_acct_dept, _
															I3_gl_dt, _
															I4_gl_input_type, _
															IG1_a_batch_no) 
	
	
	If CheckSYSTEMError(Err, True) = True Then		
       Set PAGG116_cAMngOneBtchToGlSvr = Nothing
       Exit Sub
    End If
    
    Set PAGG116_cAMngOneBtchToGlSvr  = Nothing

	Response.Write " <Script Language=vbscript>			" & vbCr
	Response.Write " With parent						" & vbCr
	Response.Write " .DbSaveOk   " & vbCr
	Response.Write " End With   " & vbCr
	Response.Write " </Script>  " & vbCr              

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
