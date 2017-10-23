<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

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

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    iKey1 = iKey1 & " AND a.bonus_yymm = " & FilterVar(Replace(lgKeyStream(1),gComDateType,""), "''", "S")
    iKey1 = iKey1 & " AND a.emp_no LIKE  " & FilterVar(lgKeyStream(2) & "%", "''", "S") & ""
    iKey1 = iKey1 & " AND a.emp_no = b.emp_no"
    iKey1 = iKey1 & " AND b.internal_cd LIKE  " & FilterVar(lgKeyStream(3) & "%", "''", "S") & ""
    
    Call SubMakeSQLStatements("MC",iKey1,"X",C_EQ)                                 '¡Ù : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("bonus_rate"),    2,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("add_rate"),    2,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("minus1_rate"),    2,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("minus2_rate"),    2,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("prov_rate"),    2,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("prov_amt"),    ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("grand_amt"),    ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("grand_rate"),    2,0)

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

	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
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
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "INSERT INTO HEA030T("
    lgStrSQL = lgStrSQL & " BONUS_TYPE,"
    lgStrSQL = lgStrSQL & " BONUS_YYMM,"
    lgStrSQL = lgStrSQL & " EMP_NO,"
    lgStrSQL = lgStrSQL & " BONUS_RATE,"
    lgStrSQL = lgStrSQL & " ADD_RATE,"
    lgStrSQL = lgStrSQL & " MINUS1_RATE,"
    lgStrSQL = lgStrSQL & " MINUS2_RATE,"
    lgStrSQL = lgStrSQL & " PROV_RATE,"
    lgStrSQL = lgStrSQL & " PROV_AMT,"
    lgStrSQL = lgStrSQL & " GRAND_AMT,"
    lgStrSQL = lgStrSQL & " GRAND_RATE,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Replace(arrColVal(3),gComDateType,""), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S") & ","

    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(8),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(11),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(12),0) & ","

    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
    
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

    lgStrSQL = "UPDATE  HEA030T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " BONUS_RATE = " & UNIConvNum(arrColVal(5),0) & ","
    lgStrSQL = lgStrSQL & " ADD_RATE = " & UNIConvNum(arrColVal(6),0) & ","
    lgStrSQL = lgStrSQL & " MINUS1_RATE = " & UNIConvNum(arrColVal(7),0) & ","
    lgStrSQL = lgStrSQL & " MINUS2_RATE = " & UNIConvNum(arrColVal(8),0) & ","
    lgStrSQL = lgStrSQL & " PROV_RATE = " & UNIConvNum(arrColVal(9),0) & ","
    lgStrSQL = lgStrSQL & " PROV_AMT = " & UNIConvNum(arrColVal(10),0) & ","
    lgStrSQL = lgStrSQL & " GRAND_AMT = " & UNIConvNum(arrColVal(11),0) & ","
    lgStrSQL = lgStrSQL & " GRAND_RATE = " & UNIConvNum(arrColVal(12),0) & ","

    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "       BONUS_TYPE = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND BONUS_YYMM = " & FilterVar(Replace(arrColVal(3),gComDateType,""), "''", "S")
    lgStrSQL = lgStrSQL & "   AND EMP_NO = " & FilterVar(arrColVal(4), "''", "S")
    
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

    lgStrSQL = "DELETE  HEA030T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "       BONUS_TYPE = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND BONUS_YYMM = " & FilterVar(Replace(arrColVal(3),gComDateType,""), "''", "S")
    lgStrSQL = lgStrSQL & "   AND EMP_NO = " & FilterVar(arrColVal(4), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
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
               Case "C"
                       lgStrSQL = "Select top " & iSelCount & " b.emp_no, b.name, bonus_rate, add_rate, minus1_rate,  minus2_rate, prov_rate, prov_amt, grand_amt, grand_rate " 
                       lgStrSQL = lgStrSQL & " From  HEA030T a, HAA010T b "
                       lgStrSQL = lgStrSQL & " Where a.bonus_type " & pComp & pCode
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
                .ggoSpread.Source  = .frm1.vspdData
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
