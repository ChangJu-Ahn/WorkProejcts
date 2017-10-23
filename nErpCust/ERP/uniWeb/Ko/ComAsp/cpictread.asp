<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
%>
<!-- #Include file="../inc/IncServer.asp" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/lgsvrvariables.inc" -->
<!-- #Include file="../inc/incServeradodb.asp" -->
<%
    Dim lgtxtKeyValue
    Dim lgtxtDKeyValue
    Dim lgtxtTable
    Dim lgtxtField
    Dim lgtxtKey
    
	'------ Developer Coding part (Start)  -----------
    
    
    
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    
    lgtxtKeyValue     = Request("txtKeyValue")                                       '��: Read Key value
    lgtxtDKeyValue    = Request("txtDKeyValue")                                      '��: Read Default Key value
    lgtxtTable        = Request("txtTable")                                          '��: Read Table Name
    lgtxtField        = Request("txtField")                                          '��: Read Field
    lgtxtKey          = Request("txtKey")                                            '��: Read Key

	'------ Developer Coding part (Start)  -----------  
    
    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection

    Call SubBizQuery()

    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iProcessOk
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    iProcessOk = False
    lgStrSQL =            "Select " & lgtxtField
    lgStrSQL = lgStrSQL & " From  " & lgtxtTable
    lgStrSQL = lgStrSQL & " Where " & lgtxtKey & "=" & FilterVar(lgtxtKeyValue,"''","S")

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       Call SubCloseRs(lgObjRs)   
       lgStrSQL =            "Select " & lgtxtField
       lgStrSQL = lgStrSQL & " From  " & lgtxtTable
       lgStrSQL = lgStrSQL & " Where " & lgtxtKey & "=" & LCase(FilterVar(lgtxtDKeyValue,"''","S"))

       If  FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then 
       Else
           iProcessOk = True  
       End If
    Else
       iProcessOk = True   
    End If   
    If iProcessOk = True Then       
        Response.Buffer = true
        Response.Clear
        Response.BinaryWrite lgObjRs(0)
    End If    
    Call SubCloseRs(lgObjRs)                                                    '�� : Release RecordSSet
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
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case pOpCode
        Case "MC"
        Case "MD"
        Case "MR"
        Case "MU"
    End Select
End Sub
%>

