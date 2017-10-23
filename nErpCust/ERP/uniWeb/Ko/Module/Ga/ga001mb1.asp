<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->

<%

    On Error Resume Next
    Err.Clear                                                                        '¢Ð: Clear Error status
    
	Call LoadBasisGlobalInf()

    Dim txtMinor
    Dim txtCost

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    dim lginRate
    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '¢Ð: "P"(Prev search) "N"(Next search)

'    Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
'   lgMaxCount        = CInt(Request("lgMaxCount"))                                  '¢Ð: Fetch count at a time for VspdData
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
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)
             'Call SubBizSaveMulti()
            ' CALL SubBizSaveMultiDelete()
             'Call bulk_disposal()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim txtGlNo
    Dim iLcNo

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

																					 '¢Ð : Release RecordSSet
    Call SubBizQueryMulti()

End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data
'============================================================================================================
Sub SubBizSave()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '¢Ð: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '¢Ð: Create
              Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE                                                             '¢Ð: Update
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

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    dim pKey1
    Dim var1
    Dim strSql
    Dim strWhere

   On Error Resume Next                                                             '¢Ð: Protect system from crashing
   Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    strSql =          "  SELECT  distinct major_cd "
    strSql = strSql & "    FROM  b_minor "
    strSql = strSql & "   WHERE  major_cd not in ( select major_cd from g_option ) "
    strSql = strSql & "     AND  major_cd in (" & FilterVar("G1009", "''", "S") & "," & FilterVar("G1010", "''", "S") & "," & FilterVar("G1011", "''", "S") & "," & FilterVar("G1012", "''", "S") & "," & FilterVar("G1022", "''", "S") & "," & FilterVar("G1024", "''", "S") & ") "

    If 	FncOpenRs("R",lgObjConn,lgObjRs,strSql,"X","X") = False Then
    Else
        Do While Not lgObjRs.EOF
            var1 = lgObjRs("MAJOR_CD")
            Call SubBizSaveMultiCreate(var1)
            lgObjRs.MoveNext
        Loop
    End If


    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                                   '¡Ù: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					'¢Ð : No data is found.
        Call SetErrorStatus()

    Else
        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""

        iDx = 1

        Do While Not lgObjRs.EOF

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))
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

    Call SubHandleError("MR",lgObjRs,Err)

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

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

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data

        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create

            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
        End Select

        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If

    Next

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    Dim txtGlNo

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------


    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(var1)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    IF UCase(Trim(var1)) = "G1024" Then
		lgStrSQL = " INSERT INTO G_OPTION ("
		lgStrSQL = lgStrSQL & " MAJOR_CD , "
		lgStrSQL = lgStrSQL & " MINOR_CD , "
		lgStrSQL = lgStrSQL & " INSRT_USER_ID,"
		lgStrSQL = lgStrSQL & " INSRT_DT,"
		lgStrSQL = lgStrSQL & " UPDT_USER_ID,"
		lgStrSQL = lgStrSQL & " UPDT_DT"
		lgStrSQL = lgStrSQL & " ) VALUES ( "
		lgStrSQL = lgStrSQL & FilterVar(var1, "''", "S") & ","
		lgStrSQL = lgStrSQL & " " & FilterVar("ST", "''", "S") & ","
		lgStrSQL = lgStrSQL & FilterVar(gUsrID, "''", "S")                        & ","
		lgStrSQL = lgStrSQL & "getdate(),"
		lgStrSQL = lgStrSQL & FilterVar(gUsrID, "''", "S")                        & ","
		lgStrSQL = lgStrSQL & "getdate())"
    
	ELSE
		lgStrSQL = " INSERT INTO G_OPTION ("
		lgStrSQL = lgStrSQL & " MAJOR_CD , "
		lgStrSQL = lgStrSQL & " MINOR_CD , "
		lgStrSQL = lgStrSQL & " INSRT_USER_ID,"
		lgStrSQL = lgStrSQL & " INSRT_DT,"
		lgStrSQL = lgStrSQL & " UPDT_USER_ID,"
		lgStrSQL = lgStrSQL & " UPDT_DT"
		lgStrSQL = lgStrSQL & " ) VALUES ( "
		lgStrSQL = lgStrSQL & FilterVar(var1, "''", "S") & ","
		lgStrSQL = lgStrSQL & " " & FilterVar("1", "''", "S") & " ,"
		lgStrSQL = lgStrSQL & FilterVar(gUsrID, "''", "S")                        & ","
		lgStrSQL = lgStrSQL & "getdate(),"
		lgStrSQL = lgStrSQL & FilterVar(gUsrID, "''", "S")                        & ","
		lgStrSQL = lgStrSQL & "getdate())"
   
    END IF

     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
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
    lgStrSQL = "UPDATE  G_OPTION"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " MINOR_CD      = " & FilterVar(UCase(arrColVal(3)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID  = " & FilterVar(gUsrID, "''", "S")                & ","
    lgStrSQL = lgStrSQL & " UPDT_DT       = getdate()"
    lgStrSQL = lgStrSQL & " WHERE  "
    lgStrSQL = lgStrSQL & " MAJOR_CD  = "&FilterVar(UCase(arrColVal(2)), "''", "S")
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
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
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount


    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
                        Select Case  lgPrevNext
                             Case ""

                             Case "P"

                             Case "N"

                        End Select
               Case "D"

               Case "U"

               Case "C"
            end select
        Case "M"
           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
           Select Case Mid(pDataType,2,1)
               Case "C"

               Case "D"

               Case "R"
                       lgStrSQL = "Select Top " &iSelCount& " B.MAJOR_CD, A.MAJOR_NM ,B.MINOR_CD,C.MINOR_NM"
                       lgStrSQL = lgStrSQL & " From B_MAJOR A, G_OPTION B , B_MINOR C"
                       lgStrSQL = lgStrSQL & " WHERE A.MAJOR_CD = B.MAJOR_CD"
                       lgStrSQL = lgStrSQL & "  and  C.MAJOR_CD = A.MAJOR_CD"
                       lgStrSQL = lgStrSQL & "  and  B.MINOR_CD = C.MINOR_CD "
                       lgStrSQL = lgStrSQL & "order by b.major_cd "
               Case "U"

               Case "B"

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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"                               '¢Ð : Display data
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .DBQueryOk
	         End with
          End If
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If
    End Select

</Script>

