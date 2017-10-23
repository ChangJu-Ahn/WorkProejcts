<%@ LANGUAGE=VBSCript%>
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
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

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

    If lgKeyStream(0) = "" then
    Else
       strWhere = FilterVar(lgKeyStream(0), "''", "S")
    End if
    
    If lgKeyStream(1) = "" then
       strWhere = strWhere & " AND PAY_CD LIKE " & "" & FilterVar("%", "''", "S") & ""
    Else
       strWhere = strWhere & " AND PAY_CD LIKE " & FilterVar(lgKeyStream(1), "''", "S")
    End if 
    
    Call SubMakeSQLStatements("MR",strWhere,"X","=")                                 '¡Ù : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1

        Do While Not lgObjRs.EOF
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_TYPE_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EXCEPT_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EXCEPT_TYPE_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_DAY"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OVER_UNDER"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OVER_UNDER_nm"))
            
            If ConvSPChars(lgObjRs("PROV_FLAG")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "YES"
            Else
                lgstrData = lgstrData & Chr(11) & "NO"
            End if
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(CDbl(lgObjRs("PROV_RATE")) * 100, 2,0)   
            If ConvSPChars(lgObjRs("CALCU_TYPE")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "YES"
            Else
                lgstrData = lgstrData & Chr(11) & "NO"
            End if

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
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    if arrColVal(8)="YES" then
        arrColVal(8) = "Y"
    Else
        arrColVal(8) = "N"
    End if
    if arrColVal(10)="YES" then
        arrColVal(10) = "Y"
    Else
        arrColVal(10) = "N"
    End if    

	arrColVal(9) = UNIConvNum(arrColVal(9), 0)/100
    arrColVal(11) = UNIConvNum(arrColVal(11), 0)/100
    
    lgStrSQL = "INSERT INTO HDA030T( PAY_CD, ALLOW_CD, EMP_TYPE," 
    lgStrSQL = lgStrSQL & " EXCEPT_TYPE, PROV_FLAG,PROV_RATE, CALCU_TYPE,WK_DAY," 
    lgStrSQL = lgStrSQL & " MINI_AMT, OVER_UNDER," 
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO , ISRT_DT  , UPDT_EMP_NO , UPDT_DT )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(8)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(9))), "", "D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(10)), "''", "S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(6))), "", "D")     & ","

    lgStrSQL = lgStrSQL & "0 ,"
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")    & ","
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

    if arrColVal(8)="YES" then
        arrColVal(8) = "Y"
    Else
        arrColVal(8) = "N"
    End if
    if arrColVal(10)="YES" then
        arrColVal(10) = "Y"
    Else
        arrColVal(10) = "N"
    End if
    arrColVal(9) = UNIConvNum(arrColVal(9), 0)/100
    arrColVal(11) = UNIConvNum(arrColVal(11), 0)/100

    lgStrSQL = "UPDATE  HDA030T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " WK_DAY   = " & FilterVar(Trim(UCase(arrColVal(6))), "", "D")   & ","
    lgStrSQL = lgStrSQL & " PROV_FLAG   = " & FilterVar(UCase(arrColVal(8)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " PROV_RATE  = " & FilterVar(Trim(UCase(arrColVal(9))), "", "D")   & ","
    lgStrSQL = lgStrSQL & " CALCU_TYPE   = " & FilterVar(UCase(arrColVal(10)), "''", "S")  
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " PAY_CD   = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND ALLOW_CD  = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_TYPE  = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & " AND EXCEPT_TYPE  = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    lgStrSQL = lgStrSQL & " AND OVER_UNDER  = " & FilterVar(UCase(arrColVal(7)), "''", "S")

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

    lgStrSQL = "DELETE  HDA030T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " PAY_CD   = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND ALLOW_CD  = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_TYPE  = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & " AND EXCEPT_TYPE  = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    lgStrSQL = lgStrSQL & " AND OVER_UNDER  = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    
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
                       lgStrSQL = "SELECT TOP " & iSelCount   
                       lgStrSQL = lgStrSQL & " PAY_CD,  "
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ",PAY_CD) PAY_CD_nm,  "
                       lgStrSQL = lgStrSQL & " ALLOW_CD,  EMP_TYPE, "
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0053", "''", "S") & ",EMP_TYPE) EMP_TYPE_nm,  "
                       lgStrSQL = lgStrSQL & " EXCEPT_TYPE, "
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0054", "''", "S") & ",EXCEPT_TYPE) EXCEPT_TYPE_nm,  "
                       lgStrSQL = lgStrSQL & " PROV_FLAG,  PROV_RATE, CALCU_TYPE, "
                       lgStrSQL = lgStrSQL & " WK_DAY,  MINI_AMT,  OVER_UNDER, "                       
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0100", "''", "S") & ",OVER_UNDER) OVER_UNDER_nm  "
                       lgStrSQL = lgStrSQL & " FROM  HDA030T "
                       lgStrSQL = lgStrSQL & " WHERE ALLOW_CD " & pComp & pCode
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
