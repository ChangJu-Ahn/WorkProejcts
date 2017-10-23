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

    lgCurrentSpd      = Request("lgCurrentSpd")                                      '¢Ð: "M"(Spread #1) "S"(Spread #2)
    
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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")

    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQ)                                 '¡Ù: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        If lgCurrentSpd = "M" Then
           Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Else
           Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        End If   
        Call SetErrorStatus()
    Else
    
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
            Select Case lgCurrentSpd
               Case "M"
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD"))
                      lgstrData = lgstrData & Chr(11) & ""
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_NAME"))
                      lgstrData = lgstrData & Chr(11) & ""
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD"))
                      lgstrData = lgstrData & Chr(11) & ""
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD_NM"))
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_SEQ"))
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CALC_YN"))
               Case Else
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD"))
                      lgstrData = lgstrData & Chr(11) & ""
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_NAME"))
                      lgstrData = lgstrData & Chr(11) & ""
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD"))
                      lgstrData = lgstrData & Chr(11) & ""
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD_NM"))
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_SEQ"))
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LEND_BASE"))
                      lgstrData = lgstrData & Chr(11) & ""
                      lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CALC_YN"))
               End Select      
            
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

'============================================================================================================
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Select Case lgCurrentSpd
    Case "M"
		lgStrSQL = "INSERT INTO HDF400T("
		lgStrSQL = lgStrSQL & " PROV_TYPE    ," 
		lgStrSQL = lgStrSQL & " PAY_CD       ," 
		lgStrSQL = lgStrSQL & " ALLOW_NAME   ," 
		lgStrSQL = lgStrSQL & " ALLOW_CD     ," 
		lgStrSQL = lgStrSQL & " ALLOW_SEQ    ," 
		lgStrSQL = lgStrSQL & " CALC_YN      ," 
		lgStrSQL = lgStrSQL & " ISRT_EMP_NO  ," 
		lgStrSQL = lgStrSQL & " ISRT_DT      ," 
		lgStrSQL = lgStrSQL & " UPDT_EMP_NO  ," 
		lgStrSQL = lgStrSQL & " UPDT_DT      )" 
		lgStrSQL = lgStrSQL & " VALUES(" 
		lgStrSQL = lgStrSQL & FilterVar("1", "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "", "D")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
		lgStrSQL = lgStrSQL & ")"
    Case Else
		lgStrSQL = "INSERT INTO HDF400T("
		lgStrSQL = lgStrSQL & " PROV_TYPE    ," 
		lgStrSQL = lgStrSQL & " PAY_CD       ," 
		lgStrSQL = lgStrSQL & " ALLOW_NAME   ," 
		lgStrSQL = lgStrSQL & " ALLOW_CD     ," 
		lgStrSQL = lgStrSQL & " ALLOW_SEQ    ," 
        lgStrSQL = lgStrSQL & " LEND_BASE    ," 
		lgStrSQL = lgStrSQL & " CALC_YN      ," 
		lgStrSQL = lgStrSQL & " ISRT_EMP_NO  ," 
		lgStrSQL = lgStrSQL & " ISRT_DT      ," 
		lgStrSQL = lgStrSQL & " UPDT_EMP_NO  ," 
		lgStrSQL = lgStrSQL & " UPDT_DT      )" 
		lgStrSQL = lgStrSQL & " VALUES(" 
		lgStrSQL = lgStrSQL & FilterVar("2", "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "", "D")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(7), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
		lgStrSQL = lgStrSQL & ")"
	End Select
	
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

    Select Case lgCurrentSpd
	Case "M"
	    lgStrSQL = "UPDATE  HDF400T"
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & "     PAY_CD      = " & FilterVar(UCase(arrColVal(2)), "''", "S")   & ","
		lgStrSQL = lgStrSQL & "     ALLOW_NAME  = " & FilterVar(UCase(arrColVal(3)), "''", "S")   & ","
		lgStrSQL = lgStrSQL & "     ALLOW_CD    = " & FilterVar(UCase(arrColVal(4)), "''", "S")   & ","
		lgStrSQL = lgStrSQL & "     ALLOW_SEQ   = " & FilterVar(Trim(UCase(arrColVal(5))), "", "D")   & ","
		lgStrSQL = lgStrSQL & "     CALC_YN     = " & FilterVar(UCase(arrColVal(6)), "''", "S")   & ","
		lgStrSQL = lgStrSQL & "     UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")                      & ","
		lgStrSQL = lgStrSQL & "     UPDT_DT     = " & FilterVar(GetSvrDateTime, "''", "S")
		lgStrSQL = lgStrSQL & " WHERE PROV_TYPE  = " & FilterVar("1", "''", "S") & "  " 
        lgStrSQL = lgStrSQL & "   AND PAY_CD     = " & FilterVar(UCase(arrColVal(2)), "''", "S")
        lgStrSQL = lgStrSQL & "   AND ALLOW_NAME = " & FilterVar(UCase(arrColVal(3)), "''", "S")
        lgStrSQL = lgStrSQL & "   AND ALLOW_CD   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    Case Else
	    lgStrSQL = "UPDATE  HDF400T"
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & "     PAY_CD      = " & FilterVar(UCase(arrColVal(2)), "''", "S")   & ","
		lgStrSQL = lgStrSQL & "     ALLOW_NAME  = " & FilterVar(UCase(arrColVal(3)), "''", "S")   & ","
		lgStrSQL = lgStrSQL & "     ALLOW_CD    = " & FilterVar(UCase(arrColVal(4)), "''", "S")   & ","
		lgStrSQL = lgStrSQL & "     ALLOW_SEQ   = " & FilterVar(Trim(UCase(arrColVal(5))), "", "D")   & ","
        lgStrSQL = lgStrSQL & "     LEND_BASE   = " & FilterVar(UCase(arrColVal(6)), "''", "S")   & ","
		lgStrSQL = lgStrSQL & "     CALC_YN     = " & FilterVar(UCase(arrColVal(7)), "''", "S")   & ","
		lgStrSQL = lgStrSQL & "     UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")                      & ","
		lgStrSQL = lgStrSQL & "     UPDT_DT     = " & FilterVar(GetSvrDateTime, "''", "S")
		lgStrSQL = lgStrSQL & " WHERE PROV_TYPE  = " & FilterVar("2", "''", "S") & " " 
        lgStrSQL = lgStrSQL & "   AND PAY_CD     = " & FilterVar(UCase(arrColVal(2)), "''", "S")
        lgStrSQL = lgStrSQL & "   AND ALLOW_NAME = " & FilterVar(UCase(arrColVal(3)), "''", "S")
        lgStrSQL = lgStrSQL & "   AND ALLOW_CD   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
	End Select

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

    Select Case lgCurrentSpd
	Case "M"
		lgStrSQL =            "DELETE HDF400T"
		lgStrSQL = lgStrSQL & " WHERE PROV_TYPE  = " & FilterVar("1", "''", "S") & "  " 
        lgStrSQL = lgStrSQL & "   AND PAY_CD     = " & FilterVar(UCase(arrColVal(2)), "''", "S")
        lgStrSQL = lgStrSQL & "   AND ALLOW_NAME = " & FilterVar(UCase(arrColVal(3)), "''", "S")
        lgStrSQL = lgStrSQL & "   AND ALLOW_CD   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    Case Else
		lgStrSQL =            "DELETE HDF400T"
		lgStrSQL = lgStrSQL & " WHERE PROV_TYPE  = " & FilterVar("2", "''", "S") & " " 
        lgStrSQL = lgStrSQL & "   AND PAY_CD     = " & FilterVar(UCase(arrColVal(2)), "''", "S")
        lgStrSQL = lgStrSQL & "   AND ALLOW_NAME = " & FilterVar(UCase(arrColVal(3)), "''", "S")
        lgStrSQL = lgStrSQL & "   AND ALLOW_CD   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
	End Select
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
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
    
    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       If lgCurrentSpd = "M" Then
                          lgStrSQL = "Select TOP " & iSelCount  & " PAY_CD, ALLOW_NAME, ALLOW_SEQ, CALC_YN, "
                          lgStrSQL = lgStrSQL & " ALLOW_CD, dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ", ALLOW_CD, " & FilterVar("1", "''", "S") & " ) ALLOW_CD_NM "
                          lgStrSQL = lgStrSQL & " From  HDF400T "
                          lgStrSQL = lgStrSQL & " WHERE PROV_TYPE = " & FilterVar("1", "''", "S") & "  AND PAY_CD " & pComp & pCode
                          lgStrSQL = lgStrSQL & " ORDER BY ALLOW_NAME,ALLOW_CD ASC"
                       Else
                          lgStrSQL = "Select TOP " & iSelCount  & " PAY_CD, ALLOW_NAME, ALLOW_SEQ, LEND_BASE, CALC_YN, "
                          lgStrSQL = lgStrSQL & " ALLOW_CD, dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ", ALLOW_CD, " & FilterVar("2", "''", "S") & ") ALLOW_CD_NM "
                          lgStrSQL = lgStrSQL & " From  HDF400T "
                          lgStrSQL = lgStrSQL & " WHERE PROV_TYPE = " & FilterVar("2", "''", "S") & " AND PAY_CD " & pComp & pCode
                          lgStrSQL = lgStrSQL & " ORDER BY ALLOW_NAME,ALLOW_CD ASC"
                       End If             
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
    On Error Resume Next                                                              '¢Ð: Protect system from crashing
    Err.Clear                                                                         '¢Ð: Clear Error status

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
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
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
                If  Trim("<%=lgCurrentSpd%>") = "M" Then
                   .ggoSpread.Source     = .frm1.vspdData
                   .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                Else
                   .ggoSpread.Source     = .frm1.vspdData2
                   .lgStrPrevKey1    = "<%=lgStrPrevKey%>"
                End If
                
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   

       Case "<%=UID_M0003%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	
