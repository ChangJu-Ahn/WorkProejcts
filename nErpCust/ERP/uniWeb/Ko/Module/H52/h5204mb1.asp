<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
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
    Call SubBizQueryMulti()
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
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    strWhere = iKey1
    strWhere = strWhere & " And hdd010t.emp_no = haa010t.emp_no "
    
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                              '¡Ù : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '¢Ð : No data is found.
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""

        iDx       = 1
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BORW_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BORW_NM"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INTREST_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INTREST_TYPE_nm"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("BORW_DT"),Null)
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("EXPIR_DT"),Null)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RESRV_DUR"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BORW_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_INTCHNG_AVG"), ggAmtOfMoney.DecPoint,0)                        
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TOT_INCHNG_CNT"), 0,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_INTCHNG"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BONUS_INTCHNG_CNT"), 0,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BONUS_INTCHNG"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("INTREST_RATE"), 2,0)

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If

        Loop
    End If

    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
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
'SetErrorStatus()
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "INSERT INTO HDD010T         ("
    lgStrSQL = lgStrSQL & " EMP_NO ,"
    lgStrSQL = lgStrSQL & " BORW_CD ,"
    lgStrSQL = lgStrSQL & " INTREST_TYPE ,"
    lgStrSQL = lgStrSQL & " BORW_DT ,"
    lgStrSQL = lgStrSQL & " EXPIR_DT ,"
    lgStrSQL = lgStrSQL & " RESRV_DUR ,"
    lgStrSQL = lgStrSQL & " BORW_TOT_AMT ,"
    lgStrSQL = lgStrSQL & " PAY_INTCHNG_AVG ,"
    lgStrSQL = lgStrSQL & " TOT_INCHNG_CNT ,"
    lgStrSQL = lgStrSQL & " PAY_INTCHNG ,"
    lgStrSQL = lgStrSQL & " BONUS_INTCHNG_CNT ,"
    lgStrSQL = lgStrSQL & " BONUS_INTCHNG ,"
    lgStrSQL = lgStrSQL & " INTREST_RATE ,"
    lgStrSQL = lgStrSQL & " Isrt_emp_no ,"
    lgStrSQL = lgStrSQL & " isrt_dt ,"
    lgStrSQL = lgStrSQL & " updt_emp_no ,"
    lgStrSQL = lgStrSQL & " updt_dt )"
    lgStrSQL = lgStrSQL & " VALUES ("
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NULL),"NULL","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(6),NULL),"NULL","S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(8),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0)     & ","
   	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10),0)     & ","
   	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(11),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(12),0)     & ","
   	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(13),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(14),0)     & ","

    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & ","
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & ","
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "UPDATE  HDD010T"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " EXPIR_DT          = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(6),NULL),"NULL","S")     & ","
    lgStrSQL = lgStrSQL & " RESRV_DUR         = " & UNIConvNum(arrColVal(7),0)     & ","
   	lgStrSQL = lgStrSQL & " BORW_TOT_AMT      = " & UNIConvNum(arrColVal(8),0)     & ","
    lgStrSQL = lgStrSQL & " PAY_INTCHNG_AVG   = " & UNIConvNum(arrColVal(9),0)     & ","
    lgStrSQL = lgStrSQL & " TOT_INCHNG_CNT    = " & UNIConvNum(arrColVal(10),0)     & ","
    lgStrSQL = lgStrSQL & " PAY_INTCHNG       = " & UNIConvNum(arrColVal(11),0)     & ","
   	lgStrSQL = lgStrSQL & " BONUS_INTCHNG_CNT = " & UNIConvNum(arrColVal(12),0)     & ","
   	lgStrSQL = lgStrSQL & " BONUS_INTCHNG     = " & UNIConvNum(arrColVal(13),0)     & ","
    lgStrSQL = lgStrSQL & " INTREST_RATE      = " & UNIConvNum(arrColVal(14),0)     & ","
    lgStrSQL = lgStrSQL & " updt_emp_no       = " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " updt_dt           = " & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE EMP_NO       = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND BORW_CD      = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND INTREST_TYPE = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND BORW_DT      = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NULL),"NULL","S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "DELETE  HDD010T "
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " EMP_NO       = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " And BORW_CD  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " And INTREST_TYPE  = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & " And BORW_DT       = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NULL),"NULL","S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
       Case "M"

           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount
                       lgStrSQL = lgStrSQL & " hdd010t.borw_cd , dbo.ufn_GetCodeName(" & FilterVar("H0043", "''", "S") & ",hdd010t.borw_cd) borw_nm , "
                       lgStrSQL = lgStrSQL & " hdd010t.intrest_type , dbo.ufn_GetCodeName(" & FilterVar("H0044", "''", "S") & ",hdd010t.intrest_type) intrest_type_nm , "
                       lgStrSQL = lgStrSQL & " hdd010t.borw_dt , hdd010t.expir_dt , "
                       lgStrSQL = lgStrSQL & " hdd010t.resrv_dur ,hdd010t.borw_tot_amt ,hdd010t.tot_inchng_cnt ,hdd010t.pay_intchng , "
                       lgStrSQL = lgStrSQL & " hdd010t.pay_intchng_avg , hdd010t.bonus_intchng_cnt , "
                       lgStrSQL = lgStrSQL & " hdd010t.bonus_intchng ,hdd010t.intrest_rate, "
                       lgStrSQL = lgStrSQL & " haa010t.emp_no ,haa010t.name  "
                       lgStrSQL = lgStrSQL & " From   haa010t , hdd010t "
                       lgStrSQL = lgStrSQL & " Where hdd010t.emp_no " & pComp & pCode
           End Select
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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk
	         End with
          End If
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If
       Case "<%=UID_M0002%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else
          End If
    End Select


</Script>
