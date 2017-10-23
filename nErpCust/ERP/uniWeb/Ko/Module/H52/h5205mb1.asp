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

	dim lgGetSvrDateTime
	
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	lgGetSvrDateTime = GetSvrDateTime
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim strIntchng_yymm_dt
    Dim strBorw_cd
    Dim strEmp_no
    Dim strIntrest_type
    Dim strInternal_cd
    Dim strWhere
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strIntchng_yymm_dt = FilterVar(lgKeyStream(0), "''", "S") 'YY-MM을 YYMM으로 
    strBorw_cd         = FilterVar(Trim(UCase(lgKeyStream(1))),"'%'", "S")
    strEmp_no          = FilterVar(Trim(UCase(lgKeyStream(2))),"'%'", "S")
    strIntrest_type    = FilterVar(Trim(UCase(lgKeyStream(3))),"'%'", "S")
    strInternal_cd     = "" & FilterVar("%", "''", "S") & ""              ' 자료권한이 생기면 ida.auth_internal_cd로 한다.

    strWhere = strIntchng_yymm_dt
    strWhere = strWhere & " And ( hdd020t.emp_no = haa010t.emp_no ) "
    strWhere = strWhere & " And ( haa010t.internal_cd LIKE " & strInternal_cd & ") "
    strWhere = strWhere & " And ( hdd020t.borw_cd LIKE " & strBorw_cd & ") "
    strWhere = strWhere & " And ( hdd020t.intrest_type LIKE " & strIntrest_type & ") "
    strWhere = strWhere & " And ( hdd020t.emp_no LIKE " & strEmp_no & ")  "
    strWhere = strWhere & " And   haa010t.internal_cd Like  " & FilterVar(lgKeyStream(4) & "%", "''", "S") & ""
    strWhere = strWhere & " Order by hdd020t.borw_cd asc,hdd020t.borw_dt asc   "

    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                              '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '☜ : No data is found.
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""

        iDx       = 1
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("borw_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("borw_NM"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("borw_dt"),Null)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("intrest_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("intrest_NM"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("pay_intchng"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("intrest_amt"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("bonus_intchng"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("tot_intchng_amt"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("borw_baln"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("borw_tot_amt"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("pay_intchng_cnt"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("bonus_intchng_cnt"), ggAmtOfMoney.DecPoint,0)

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
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data

        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HDD020T         ("
    lgStrSQL = lgStrSQL & " intchng_yymm         ,"
    lgStrSQL = lgStrSQL & " emp_no         ,"
    lgStrSQL = lgStrSQL & " borw_cd         ,"
    lgStrSQL = lgStrSQL & " borw_dt         ,"
    lgStrSQL = lgStrSQL & " intrest_type         ,"

    lgStrSQL = lgStrSQL & " pay_intchng         ,"
    lgStrSQL = lgStrSQL & " intrest_amt         ,"
    lgStrSQL = lgStrSQL & " bonus_intchng         ,"
    lgStrSQL = lgStrSQL & " tot_intchng_amt         ,"
    lgStrSQL = lgStrSQL & " borw_baln         ,"
    lgStrSQL = lgStrSQL & " borw_tot_amt         ,"
    lgStrSQL = lgStrSQL & " pay_intchng_cnt         ,"
    lgStrSQL = lgStrSQL & " bonus_intchng_cnt         ,"

    lgStrSQL = lgStrSQL & " Isrt_emp_no         ,"
    lgStrSQL = lgStrSQL & " isrt_dt         ,"
    lgStrSQL = lgStrSQL & " updt_emp_no     ,"
    lgStrSQL = lgStrSQL & " updt_dt         )"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Replace(Trim(arrColVal(2)),gComDateType,""), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NULL),"NULL","S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0)     & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(8),0)     & ","
   	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0)     & ","
   	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10),0)     & ","
   	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(11),0)     & ","
   	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(12),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(13),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(14),0)     & ","

    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & ","
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & ","
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HDD020T"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " pay_intchng           = " & UNIConvNum(arrColVal(7),0)     & ","
    lgStrSQL = lgStrSQL & " intrest_amt           = " & UNIConvNum(arrColVal(8),0)     & ","
    lgStrSQL = lgStrSQL & " bonus_intchng         = " & UNIConvNum(arrColVal(9),0)     & ","
    lgStrSQL = lgStrSQL & " tot_intchng_amt       = " & UNIConvNum(arrColVal(10),0)     & ","
    lgStrSQL = lgStrSQL & " borw_baln             = " & UNIConvNum(arrColVal(11),0)     & ","
    lgStrSQL = lgStrSQL & " borw_tot_amt          = " & UNIConvNum(arrColVal(12),0)     & ","
    lgStrSQL = lgStrSQL & " pay_intchng_cnt       = " & UNIConvNum(arrColVal(13),0)     & ","
    lgStrSQL = lgStrSQL & " bonus_intchng_cnt     = " & UNIConvNum(arrColVal(14),0)     & ","

    lgStrSQL = lgStrSQL & " updt_emp_no           = " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " updt_dt               = " & FilterVar(lgGetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " intchng_yymm          = " & FilterVar(Replace(Trim(arrColVal(2)),gComDateType,""), "''", "S")
    lgStrSQL = lgStrSQL & " And emp_no            = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " And borw_cd           = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & " And borw_dt           = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NULL),"NULL","S")
    lgStrSQL = lgStrSQL & " And intrest_type      = " & FilterVar(UCase(arrColVal(6)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HDD020T "
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " intchng_yymm          = " & FilterVar(Replace(Trim(arrColVal(2)),gComDateType,""), "''", "S")
    lgStrSQL = lgStrSQL & " And emp_no            = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " And borw_cd           = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & " And borw_dt           = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(5),NULL),"NULL","S")
    lgStrSQL = lgStrSQL & " And intrest_type      = " & FilterVar(UCase(arrColVal(6)), "''", "S")

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
                       lgStrSQL = lgStrSQL & " haa010t.name,hdd020t.intchng_yymm,hdd020t.emp_no,hdd020t.borw_cd,dbo.ufn_GetCodeName(" & FilterVar("H0043", "''", "S") & ", hdd020t.borw_cd) borw_nm, hdd020t.borw_dt, "
                       lgStrSQL = lgStrSQL & " hdd020t.intrest_type, dbo.ufn_GetCodeName(" & FilterVar("h0044", "''", "S") & ",intrest_type) intrest_NM, hdd020t.pay_intchng,hdd020t.intrest_amt,hdd020t.bonus_intchng, "
                       lgStrSQL = lgStrSQL & " hdd020t.pay_intchng_cnt,hdd020t.bonus_intchng_cnt,hdd020t.tot_intchng_amt,hdd020t.borw_baln, "
                       lgStrSQL = lgStrSQL & " hdd020t.borw_tot_amt,hdd020t.isrt_emp_no,hdd020t.isrt_dt,hdd020t.updt_emp_no,hdd020t.updt_dt "
                       lgStrSQL = lgStrSQL & " From   haa010t,hdd020t "
                       lgStrSQL = lgStrSQL & " Where hdd020t.intchng_yymm " & pComp & pCode
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk
	         End with
          End If
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else
          End If
    End Select


</Script>
