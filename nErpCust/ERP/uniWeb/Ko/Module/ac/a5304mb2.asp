<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
    Dim lgStrPrevKey
    On Error Resume Next
    Err.Clear

    Call HideStatusWnd                                                     '☜: Hide Processing message
	Dim strYear

	Call LoadBasisGlobalInf()

    lgErrorStatus  = ""
    lgKeyStream    = Split(Request("txtKeyStream"),gColSep) 
    lgStrPrevKey   = Trim(Request("lgStrPrevKey"))                   '☜: Next Key
    strYear   = Trim(Request("txtYear"))                   '☜: Next Key

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

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
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMulti()
    Dim lgStrSQL
    On Error Resume Next
    Err.Clear

'이미 생성된 자료가 있습니다.
'해당년도 데이터가 있을경우 
    lgStrSQL = ""
    lgStrSQL = " SELECT  F_ACCT FROM A_ACCT_CLOSE_TRANSFER " & vbcr
    lgStrSQL = lgStrSQL & " WHERE YYYY			= substring(convert(varchar(8),dateadd(year,0,  convert(datetime, " & FilterVar(strYear, "''", "S") & " +" & FilterVar("0101", "''", "S") & " ,112)),112),1,4)" & vbcr
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
		Call DisplayMsgBox("800064", vbInformation, "", "", I_MKSCRIPT)
		Call SubCloseRs(lgObjRs)
		Exit Sub
		response.end
	End If

'검사결과등록 데이터가 없습니다. 이전년도 데이터가 없을경우 
    lgStrSQL = ""
    lgStrSQL = " SELECT  F_ACCT FROM A_ACCT_CLOSE_TRANSFER " & vbcr
    lgStrSQL = lgStrSQL & " WHERE YYYY			= substring(convert(varchar(8),dateadd(year,-1,  convert(datetime, " & FilterVar(strYear, "''", "S") & " +" & FilterVar("0101", "''", "S") & " ,112)),112),1,4)" & vbcr
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		Call DisplayMsgBox("174128", vbInformation, "", "", I_MKSCRIPT)
		Call SubCloseRs(lgObjRs)
		Exit Sub
		response.end
	End If



    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = ""
    lgStrSQL = " SELECT  F_ACCT ,T_ACCT FROM A_ACCT_CLOSE_TRANSFER " & vbcr
    lgStrSQL = lgStrSQL & " WHERE YYYY			= substring(convert(varchar(8),dateadd(year,-1,  convert(datetime, " & FilterVar(strYear, "''", "S") & " +" & FilterVar("0101", "''", "S") & " ,112)),112),1,4)" & vbcr
    lgStrSQL = lgStrSQL & "  AND F_ACCT NOT IN (SELECT F_ACCT FROM A_ACCT_CLOSE_TRANSFER WHERE YYYY = " & FilterVar(strYear, "''", "S") & " )" & vbcr
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
		'If isNull(lgObjRs("ACCT_FG")) = False Then
	   	'	Call DisplayMsgBox("122302", vbInformation, "", "", I_MKSCRIPT)
	   		
	   	'	Exit Sub
	   'End If
	Else
	'Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	'Call SubCloseRs(lgObjRs)

    lgStrSQL = ""
       Do While Not lgObjRs.EOF
			lgStrSQL = lgStrSQL & "INSERT INTO A_ACCT_CLOSE_TRANSFER("
			lgStrSQL = lgStrSQL & " YYYY			,"
			lgStrSQL = lgStrSQL & " SEQ				, ACCT_FG  ,"
			lgStrSQL = lgStrSQL & " F_ACCT			, T_ACCT   , "
			lgStrSQL = lgStrSQL & " INSRT_USER_ID   , INSRT_DT , "
			lgStrSQL = lgStrSQL & " UPDT_USER_ID    , UPDT_DT  )"    '16
			lgStrSQL = lgStrSQL & " VALUES(" 
			lgStrSQL = lgStrSQL & " " & FilterVar(strYear, "''", "S") & " ,"
			lgStrSQL = lgStrSQL & "" & FilterVar("1", "''", "S") & " ,"
			lgStrSQL = lgStrSQL & "" & FilterVar("1", "''", "S") & " ,"
			lgStrSQL = lgStrSQL & " " & FilterVar(lgObjRs("F_ACCT"), "''", "S") & " ,"
			lgStrSQL = lgStrSQL & " " & FilterVar(lgObjRs("T_ACCT"), "''", "S") & " ,"
			lgStrSQL = lgStrSQL & FilterVar(gUsrID, "''", "S")                       & "," 
			lgStrSQL = lgStrSQL & " GetDate()," 
			lgStrSQL = lgStrSQL & FilterVar(gUsrID, "''", "S")                       & "," 
			lgStrSQL = lgStrSQL & " GetDate())" & vbcr

          lgObjRs.MoveNext
      Loop 
    
    'Response.Write "Create : " & lgStrSQL
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
	End If
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
    If lgErrorStatus = "" Then
       Response.Write  " <Script Language=vbscript>                                  " & vbCr
       Response.Write  "    Parent.DBSaveOk   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If

End Sub


'=======

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
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
    End Select
End Sub

%>
