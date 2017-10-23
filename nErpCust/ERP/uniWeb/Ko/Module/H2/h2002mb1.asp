<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" --> 
<!-- #Include file="../../inc/ImgUpLoad.asp" -->	                       
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
    Dim byteCount
    Dim UploadRequest
    Dim RequestBin

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    Call HideStatusWnd        
    
    byteCount = Request.TotalBytes

    RequestBin = Request.BinaryRead(byteCount)
    
    Set UploadRequest = CreateObject("Scripting.Dictionary")

    BuildUploadRequest  RequestBin
  
    lgOpModeCRUD = UploadRequest.Item("txtMode").Item("Value")

    lgErrorStatus = "NO"
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    
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
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub	

'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    lgIntFlgMode = UploadRequest.Item("txtFlgMode").Item("Value")                    '¢Ð: Read Operayion Mode (CREATE, UPDATE)
    
    lgIntFlgMode = CLng(lgIntFlgMode)
  
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '¢Ð : Create
              Call SubBizSaveSingleCreate()  
        Case  Else
              Call SubBizSaveSingleUpdate()
    End Select
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    lgStrSQL = "DELETE HAA070T WHERE EMP_NO = " & FilterVar(lgKeyStream(0), "''", "S")
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    Dim Picture
    Dim lgStrSQL1

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Picture     = UploadRequest.Item("txtPath").Item("Value")
    lgKeyStream = UploadRequest.Item("txtKeyStream").Item("Value")

    lgKeyStream = Split(lgKeyStream,gColSep)

    Call SubMakeSQLStatements("U",FilterVar(lgKeyStream(0), "''", "S"))
    
    If  FncOpenRs("U",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrSQL1 = "INSERT INTO HAA070T(emp_no,isrt_emp_no,isrt_dt,updt_emp_no,updt_dt) "
        lgStrSQL1 = lgStrSQL1 & "VALUES(" & FilterVar(lgKeyStream(0), "''", "S") & ","
        lgStrSQL1 = lgStrSQL1 & FilterVar(gUsrId, "''", "S")                     & "," 
        lgStrSQL1 = lgStrSQL1 & FilterVar(GetSvrDateTime, "''", "S") & "," 
        lgStrSQL1 = lgStrSQL1 & FilterVar(gUsrId, "''", "S")                        & "," 
        lgStrSQL1 = lgStrSQL1 & FilterVar(GetSvrDateTime, "''", "S")
        lgStrSQL1 = lgStrSQL1 & ")"

        lgObjConn.Execute lgStrSQL1,,adCmdText
        Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
        call FncOpenRs("U",lgObjConn,lgObjRs,lgStrSQL,"X","X")
    End if

	lgObjRs("Photo").AppendChunk Picture
    lgObjRs.Update

    Call SubCloseRs(lgObjRs)

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("U",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Select Case pMode 
      Case "U"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From   HAA070T "
             lgStrSQL = lgStrSQL & " Where  Emp_No =  " & pCode 	
    End Select
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Select Case pOpCode
        Case "U"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
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
