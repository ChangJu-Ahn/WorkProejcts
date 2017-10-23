<%@LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<!-- #Include file="../../inc/ImgUpLoad.asp" -->	            
<%
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	Call LoadBasisGlobalInf

    Dim byteCount
    Dim UploadRequest
    Dim RequestBin
    
    Call HideStatusWnd        
    
    byteCount = Request.TotalBytes

    RequestBin = Request.BinaryRead(byteCount)
    
    Set UploadRequest = CreateObject("Scripting.Dictionary")

    BuildUploadRequest RequestBin

    lgOpModeCRUD = UploadRequest.Item("txtMode").Item("Value")
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

    lgErrorStatus = "NO"
  
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQueryz
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear
    
    Dim iIntFlgMode
    
    iIntFlgMode = UploadRequest.Item("txtFlgMode").Item("Value")                    '��: Read Operayion Mode (CREATE, UPDATE)
 
    iIntFlgMode = CLng(iIntFlgMode)
    
    Select Case iIntFlgMode
        Case  OPMD_CMODE                                                             '�� : Create
              Call SubBizSaveSingleCreate()  
        Case  Else
              Call SubBizSaveSingleUpdate()
    End Select
End Sub	
	    
'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    Dim Picture
    Dim lgStrSQL1
    Dim lgStrSQL2
    Dim intRetVal
	Dim pPB3C104
	Dim sMemo
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    Picture     = UploadRequest.Item("txtPath").Item("Value")
    sMemo     = UploadRequest.Item("txtMemo").Item("Value")
    lgKeyStream = UploadRequest.Item("txtKeyStream").Item("Value")
    lgKeyStream = Split(lgKeyStream, gColSep)

    Call SubMakeSQLStatements("U", FilterVar(lgKeyStream(0), "''", "S"))
    
    If  FncOpenRs("U", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = True Then                    'If data not exists
		lgStrSQL2 = "UPDATE Y_FAC_CAST_IMAGE "
		lgStrSQL2 = lgStrSQL2 & " SET REMRK = " & FilterVar(Trim(sMemo ),"''","S") & "," 
		lgStrSQL2 = lgStrSQL2 & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & "," 
		lgStrSQL2 = lgStrSQL2 & " UPDT_DT = GETDATE() " 
		lgStrSQL2 = lgStrSQL2 & " WHERE FAC_CAST_CD = " & FilterVar(lgKeyStream(0), "''", "S")
		lgStrSQL2 = lgStrSQL2 & "   AND GUBUN_CD = '20'"
	    
		lgObjConn.Execute lgStrSQL2,, adCmdText
		Call SubHandleError("MU", lgObjConn, lgObjRs, Err)
    Else
    
		lgStrSQL1 = "SELECT * FROM Y_CAST WHERE CAST_CD = " & FilterVar(Trim(UCase(lgKeyStream(0))),"''","S")
		If FncOpenRs("U", lgObjConn, lgObjRs, lgStrSQL1, "X", "X") = True Then 
			lgStrSQL1 = "INSERT INTO Y_FAC_CAST_IMAGE (FAC_CAST_CD, GUBUN_CD, REMRK, ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT) "
			lgStrSQL1 = lgStrSQL1 & " VALUES(" 
    		lgStrSQL1 = lgStrSQL1 & FilterVar(Trim(UCase(lgKeyStream(0))),"''","S")     & ","
			lgStrSQL1 = lgStrSQL1 & " '20', "
    		lgStrSQL1 = lgStrSQL1 & FilterVar(Trim(sMemo         ),"''","S")     & ","
			lgStrSQL1 = lgStrSQL1 & FilterVar(gUsrId, "''", "S") & ", GETDATE(), " 
			lgStrSQL1 = lgStrSQL1 & FilterVar(gUsrId, "''", "S") & ", GETDATE())" 
	
			lgObjConn.Execute lgStrSQL1,,adCmdText
			Call SubHandleError("MC", lgObjConn, lgObjRs, Err)

			Call FncOpenRs("U", lgObjConn, lgObjRs, lgStrSQL, "X", "X")
		Else
			Call SubCloseRs(lgObjRs)
			Call DisplayMsgBox("Y60040", vbInformation, "", "", I_MKSCRIPT)
			ObjectContext.SetAbort
			Call SetErrorStatus
			'Response.Write "<Script Language = VBScript>" & vbCrLf
			'Response.Write "parent.DbSaveNotOk()" & vbCrLf
			'Response.Write "</Script>" & vbCrLf
		End If
    End if

    lgObjRs("PIC_IMAGE").AppendChunk Picture
    lgObjRs.Update
    
    '----------------------------------------------------------------------------------------
		
	lgStrSQL2 = "UPDATE Y_CAST "
	lgStrSQL2 = lgStrSQL2 & " SET PIC_FLAG = " & FilterVar("Y", "''", "S") & " ,"
	lgStrSQL2 = lgStrSQL2 & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & "," 
	lgStrSQL2 = lgStrSQL2 & " UPDT_DT = GETDATE() "
	lgStrSQL2 = lgStrSQL2 & " WHERE CAST_CD = " & FilterVar(lgKeyStream(0), "''", "S")
    
	lgObjConn.Execute lgStrSQL2,, adCmdText
	Call SubHandleError("MC", lgObjConn, lgObjRs, Err)
			      
    Call SubCloseRs(lgObjRs)
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode, pCode)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

    Select Case pMode 
      Case "R"
	  Case "C"
      Case "U"
		lgStrSQL = "SELECT * FROM  Y_FAC_CAST_IMAGE WHERE FAC_CAST_CD = " & pCode & "AND GUBUN_CD = '20'"	
      Case "D"
    End Select
    
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
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode, pConn, pRs, pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case pOpCode
        Case "U"
			If CheckSYSTEMError(pErr,True) = True Then
			   Call DisplayMsgBox("122918", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
			   ObjectContext.SetAbort
			   Call SetErrorStatus
			Else
			   If CheckSQLError(pConn,True) = True Then
			      Call DisplayMsgBox("122918", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
			      ObjectContext.SetAbort
			      Call SetErrorStatus
			   End If
			End If
    End Select
End Sub
%>

<Script Language="VBScript">

    If Trim("<%=lgErrorStatus%>") = "NO" Then
       Parent.DBSaveOk
    End If
       
</Script>	

