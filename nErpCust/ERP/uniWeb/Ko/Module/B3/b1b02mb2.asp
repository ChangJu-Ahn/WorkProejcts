<%@LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<!-- #Include file="../../inc/ImgUpLoad.asp" -->	            
<%
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear
    
    Dim iIntFlgMode
    
    iIntFlgMode = UploadRequest.Item("txtFlgMode").Item("Value")                    '¢Ð: Read Operayion Mode (CREATE, UPDATE)
 
    iIntFlgMode = CLng(iIntFlgMode)
    
    Select Case iIntFlgMode
        Case  OPMD_CMODE                                                             '¢Ð : Create
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
	
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    Picture     = UploadRequest.Item("txtPath").Item("Value")
    lgKeyStream = UploadRequest.Item("txtKeyStream").Item("Value")
    lgKeyStream = Split(lgKeyStream, gColSep)

	Set pPB3C104 = Server.CreateObject("PB3C104.cBLkUpItem")    

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

	Call pPB3C104.B_LOOK_UP_ITEM(gStrGlobalCollection, lgKeyStream(0))

	If CheckSYSTEMError(Err, True) = True Then
		Set pPB3C104 = Nothing															'¢Ð: Unload Component
		Response.End
	End If

	Set pPB3C104 = Nothing															'¢Ð: Unload Component

    Call SubMakeSQLStatements("U", FilterVar(lgKeyStream(0), "''", "S"))
    
    If  FncOpenRs("U", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = True Then                    'If data not exists
		'lgObjRs("item_image").AppendChunk Picture
		'lgObjRs.Update
    Else
        lgStrSQL1 = "INSERT INTO B_ITEM_IMAGE(ITEM_CD, INSRT_USER_ID, INSRT_DT, UPDT_USER_ID, UPDT_DT) "
        lgStrSQL1 = lgStrSQL1 & "VALUES(" & FilterVar(lgKeyStream(0), "''", "S") & ", "
        lgStrSQL1 = lgStrSQL1 & FilterVar(gUsrId, "''", "S") & ", GETDATE(), " 
        lgStrSQL1 = lgStrSQL1 & FilterVar(gUsrId, "''", "S") & ", GETDATE())" 

        lgObjConn.Execute lgStrSQL1,,adCmdText
        Call SubHandleError("MC", lgObjConn, lgObjRs, Err)

		Call FncOpenRs("U", lgObjConn, lgObjRs, lgStrSQL, "X", "X")
    End if

    lgObjRs("ITEM_IMAGE").AppendChunk Picture
    lgObjRs.Update
    
    '----------------------------------------------------------------------------------------
		
	lgStrSQL2 = "UPDATE B_ITEM "
	lgStrSQL2 = lgStrSQL2 & " SET ITEM_IMAGE_FLG = " & FilterVar("Y", "''", "S") & " ,"
	lgStrSQL2 = lgStrSQL2 & " UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S") & "," 
	lgStrSQL2 = lgStrSQL2 & " UPDT_DT = GETDATE() WHERE ITEM_CD = " & FilterVar(lgKeyStream(0), "''", "S")
    
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
		lgStrSQL = "SELECT * FROM  B_ITEM_IMAGE WHERE ITEM_CD = " & pCode 	
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode, pConn, pRs, pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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

