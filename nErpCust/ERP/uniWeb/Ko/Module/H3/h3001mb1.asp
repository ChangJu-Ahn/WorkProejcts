<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
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

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    strWhere = iKey1
    strWhere = strWhere & " Order by gazet_dt desc, gazet_cd desc "
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                                 '¡Ù : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("gazet_dt"),Null)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gazet_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gazet_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sect_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sect_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("role_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("role_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("func_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("func_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gazet_resn"))
            
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

    lgStrSQL = "INSERT INTO hba010t         ("
    lgStrSQL = lgStrSQL & " emp_no , gazet_dt , gazet_cd , dept_cd , sect_cd , role_cd ," 
    lgStrSQL = lgStrSQL & " roll_pstn , pay_grd1 , pay_grd2 , func_cd , gazet_resn ," 
    lgStrSQL = lgStrSQL & " Isrt_emp_no , isrt_dt , updt_emp_no , updt_dt )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(8)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(9)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(10)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(11)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(12), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
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

    lgStrSQL = "UPDATE  hba010t"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " dept_cd          = " & FilterVar(UCase(arrColVal(5)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " sect_cd          = " & FilterVar(UCase(arrColVal(6)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " role_cd          = " & FilterVar(UCase(arrColVal(7)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " roll_pstn        = " & FilterVar(UCase(arrColVal(8)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " pay_grd1         = " & FilterVar(UCase(arrColVal(9)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " pay_grd2         = " & FilterVar(UCase(arrColVal(10)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " func_cd          = " & FilterVar(UCase(arrColVal(11)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " gazet_resn       = " & FilterVar(arrColVal(12), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " updt_emp_no      = " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " updt_dt          = " & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " emp_no           = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " And gazet_dt     = " & FilterVar(UNIConvDate(arrColVal(3)),NULL,"S")
    lgStrSQL = lgStrSQL & " And gazet_cd     = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    
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

    lgStrSQL = "DELETE  hba010t"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " emp_no          = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " And gazet_dt    = " & FilterVar(UNIConvDate(arrColVal(3)),NULL,"S")
    lgStrSQL = lgStrSQL & " And gazet_cd     = " & FilterVar(UCase(arrColVal(4)), "''", "S")

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
                       lgStrSQL = "Select TOP " & iSelCount & " emp_no    , "                       
                       lgStrSQL = lgStrSQL & "gazet_dt  , "
                       lgStrSQL = lgStrSQL & "gazet_cd  , "
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0029", "''", "S") & ",gazet_cd) gazet_nm  , "
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0035", "''", "S") & ",sect_cd) sect_nm  , "
                       lgStrSQL = lgStrSQL & "role_cd   , "					
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0026", "''", "S") & ",role_cd) role_nm  , "
                       lgStrSQL = lgStrSQL & "gazet_resn, "
                       lgStrSQL = lgStrSQL & "comp_cd   , "
                       lgStrSQL = lgStrSQL & "sect_cd   , "
                       lgStrSQL = lgStrSQL & "dept_cd   , "
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetDeptName(dept_cd,gazet_dt) dept_nm  , "                       
                       lgStrSQL = lgStrSQL & "wk_area_cd, "
                       lgStrSQL = lgStrSQL & "roll_pstn , "
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn_nm  , "
                       lgStrSQL = lgStrSQL & "pay_grd1  , "
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",pay_grd1) pay_grd1_nm  , "
                       lgStrSQL = lgStrSQL & "pay_grd2  , "
                       lgStrSQL = lgStrSQL & "func_cd   , "            
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0004", "''", "S") & ",func_cd) func_nm   "
                       lgStrSQL = lgStrSQL & " From  hba010t "
                       lgStrSQL = lgStrSQL & " Where emp_no " & pComp & pCode
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
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
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
                .ggoSpread.SSShowData "<%=lgstrData%>"                               '¢Ð : Display data                
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
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
