<%@ LANGUAGE=VBSCript %>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear 
    
    Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")                                                                           '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '¢Ð: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
	Const C_SHEETMAXROWS_D  = 100  
	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '¢Ð: Max fetched data at a time	                 

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
        Case CStr(UID_M0006)                                                         '¢Ð: Batch
'            Call SubCreateCommandObject(lgObjComm)
            Call SubBizBatch()
'            Call SubCloseCommandObject(lgObjComm)
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizbatch
' Desc : Batch
'============================================================================================================
Sub SubBizBatch()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data

        Call SubCreateCommandObject(lgObjComm)
        Call SubBizBatchMulti(arrColVal)                            '¢Ð: Run Batch
        Call SubCloseCommandObject(lgObjComm)

        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If

    Next
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim arrVal
    Dim strTemp
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    strWhere = Trim(lgKeyStream(0))
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                                 '¡Ù : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("800414", vbInformation, "Database Error", "", I_MKSCRIPT)
    Else

        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""

        iDx       = 1

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & "0"
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROGRESS_FG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_CD"))
            Select Case CStr(lgObjRs("MINOR_CD"))
                Case "19"
                    lgstrData = lgstrData & Chr(11) & "G_USP_GE020BA1"
                Case "20"
                    lgstrData = lgstrData & Chr(11) & "G_USP_GE020BA2"
                Case "21"
                    lgstrData = lgstrData & Chr(11) & "G_USP_GE020BA3"
            End Select
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(strWhere)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("NUM_OF_ERROR"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars("")
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If

        Loop
    End If

    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

End Sub

'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti(arrColVal)


    Dim IntRetCD
    Dim strMsg_cd
    Dim strYYYYMM

    strYYYYMM = Trim(lgKeyStream(0))
    With lgObjComm

        .CommandText = arrColVal(0)
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@yyyymm"     ,adVarXChar,adParamInput,6, strYYYYMM)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adVarXChar,adParamInput,13, gUsrID)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamOutput,6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@num_of_error"   ,adSmallInt ,adParamOutput,2)

        lgObjComm.Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        if  IntRetCD = 2 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT)
            lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
			Call SubHandleError("MB",lgObjComm.ActiveConnection,lgObjRs,Err)
        end if
    Else
        lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
        Call SubHandleError("MB",lgObjComm.ActiveConnection,lgObjRs,Err)
    End if
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext
                  Case " "
                  Case "P"
                  Case "N"
           End Select
        Case "M"
           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
           Select Case Mid(pDataType,2,1)
               Case "C"
               Case "D"
               Case "R"
                       lgStrSQL = " SELECT A.MINOR_CD, A.MINOR_NM, UPPER(ISNULL(C.PROGRESS_FG," & FilterVar("N", "''", "S") & " )) AS PROGRESS_FG, ISNULL(C.NUM_OF_ERROR," & FilterVar("0", "''", "S") & " ) AS NUM_OF_ERROR "
                       lgStrSQL = lgStrSQL & " FROM B_MINOR A, B_CONFIGURATION B, G_JOB_RESULT C "
                       lgStrSQL = lgStrSQL & " WHERE A.MAJOR_CD = B.MAJOR_CD "
                       lgStrSQL = lgStrSQL & " AND B.MINOR_CD *= C.JOB_CD "
                       lgStrSQL = lgStrSQL & " AND A.MINOR_CD = B.MINOR_CD "
                       lgStrSQL = lgStrSQL & " AND B.MAJOR_CD = " & FilterVar("G1007", "''", "S") & " "
                       lgStrSQL = lgStrSQL & " AND B.SEQ_NO = " & FilterVar("6", "''", "S") & " "
                       lgStrSQL = lgStrSQL & " AND B.REFERENCE = " & FilterVar("Y", "''", "S") & "  "
                       lgStrSQL = lgStrSQL & " AND C.YYYYMM " & pComp & FilterVar(pCode, "''", "S")
               Case "U"
           End Select
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
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '¢Ð: Protect system from crashing
    Err.Clear                                                                         '¢Ð: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       'Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          'Call SetErrorStatus
       End If
   End If

End Sub

%>

<Script Language="VBScript">

    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                .DBQueryOk
	          End with
          End If
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If
       Case "<%=UID_M0006%>"                                                         '¢Ð : Batch
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.ExeReflectOk
          Else
             Parent.ExeReflectNo
          End If
    End Select
</Script>
