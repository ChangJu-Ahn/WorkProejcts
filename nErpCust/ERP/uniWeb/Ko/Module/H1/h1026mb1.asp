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

    Dim lgGetSvrDateTime
    lgGetSvrDateTime = GetSvrDateTime

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")                                       '¢Ð: "P"(Prev search) "N"(Next search)
    
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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    iKey1 = FilterVar(lgKeyStream(0),"'%'", "S")
    
    if Trim(iKey1) = "" & FilterVar("%", "''", "S") & "" Then
    	Call SubMakeSQLStatements("MR",iKey1,"X",C_LIKE)
    	
    Else
    	Call SubMakeSQLStatements("MR",iKey1,"X",C_EQ)
    End if	
    
    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '¢Ð : No data is found.
        Call SetErrorStatus()
        
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
           
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_TYPE_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_TYPE"))
            
            If ConvSPChars(lgObjRs("HOLIDAY_APPLY")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "YES"
            Else
                lgstrData = lgstrData & Chr(11) & "NO"
            End if
            
            If ConvSPChars(lgObjRs("WEEK_TYPE")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "YES"
            Else
                lgstrData = lgstrData & Chr(11) & "NO"
            End if
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SAT_TIME"))
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
    
    Call SubHandleError("MR",lgObjRs,Err)
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
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    Dim iclose_dt
    Dim strPay_dt
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    IF arrColVal(4) = "YES" then
            arrColVal(4) = "Y"
    Else
            arrColVal(4) = "N"
    End IF
    IF arrColVal(5) = "YES" then
            arrColVal(5) = "Y"
    Else
            arrColVal(5) = "N"
    End IF

    lgStrSQL = "INSERT INTO HDA240T( PAY_CD, WK_TYPE, HOLIDAY_APPLY," 
    lgStrSQL = lgStrSQL & " WEEK_TYPE, SAT_TIME, " 
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(6),"0", "D") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S")
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

    IF arrColVal(4) = "YES" then
            arrColVal(4) = "Y"
    Else
            arrColVal(4) = "N"
    End IF
    IF arrColVal(5) = "YES" then
            arrColVal(5) = "Y"
    Else
            arrColVal(5) = "N"
    End IF

    lgStrSQL = "UPDATE  HDA240T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & "     HOLIDAY_APPLY = " & FilterVar(arrColVal(4),"NULL", "S") & ","
    lgStrSQL = lgStrSQL & "     WEEK_TYPE = " & FilterVar(arrColVal(5),"NULL", "S") & ","
    lgStrSQL = lgStrSQL & "     SAT_TIME = " & FilterVar(arrColVal(6),"0", "D") & ","
    lgStrSQL = lgStrSQL & "     UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & "     UPDT_DT = " & FilterVar(lgGetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE PAY_CD = " & FilterVar(arrColVal(2), "NULL", "S")
    lgStrSQL = lgStrSQL & "   AND WK_TYPE = " & FilterVar(arrColVal(3), "NULL", "S")

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

    lgStrSQL = "DELETE  HDA240T"
    lgStrSQL = lgStrSQL & " WHERE PAY_CD = " & FilterVar(arrColVal(2), "NULL", "S")
    lgStrSQL = lgStrSQL & "   AND WK_TYPE = " & FilterVar(arrColVal(3), "NULL", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements'("MR",iKey1,"X",C_EQ), ("MR",iKey1,"X",C_LIKE)
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select "
                       lgStrSQL = lgStrSQL & "		PAY_CD, "
                       lgStrSQL = lgStrSQL & "	    dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ",PAY_CD) PAY_CD_nm, "
                       lgStrSQL = lgStrSQL & "		WK_TYPE, "
                       lgStrSQL = lgStrSQL & "	    dbo.ufn_GetCodeName(" & FilterVar("H0047", "''", "S") & ",WK_TYPE) WK_TYPE_nm, "
                       lgStrSQL = lgStrSQL & "		HOLIDAY_APPLY, WEEK_TYPE, SAT_TIME "                       
                       lgStrSQL = lgStrSQL & " From  HDA240T"
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"                               '¢Ð : Display data                
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
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
