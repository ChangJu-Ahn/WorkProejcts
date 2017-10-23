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
	Const C_SHEETMAXROWS_D = 100
	Dim lgStrPrevKey
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)    

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection    
    
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
    Dim MFlag,UFlag,AFlag
    Dim strWhere
    
    MFlag = False
    UFlag = False
    AFlag = False
    strWhere = ""
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   
    If Trim(lgKeyStream(0)) <> "" Then
        MFlag = True
    End IF
    If Trim(lgKeyStream(2)) <> "" Then
        UFlag = True
    End IF
    If Trim(lgKeyStream(1)) <> "" Then
        AFlag = True
    End IF
    
    If MFlag Or UFlag Or AFlag Then        
        strWhere = " WHERE "
        If MFlag Then
            strWhere = strWhere & " MNU_ID = " & FilterVar(lgKeyStream(0), "''", "S")                                  ':MNU_ID //메뉴ID
        End If
        If UFlag Then
            IF MFlag Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " USR_ID = " & FilterVar(lgKeyStream(2), "''", "S")    ':USR_ID //사용자ID
        End If
        If AFlag Then
            IF MFlag Or UFlag Then
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " AUTH_YN = " & FilterVar(lgKeyStream(1), "''", "S")            ':AUTH_YN //권한부여여부 
        End If
    Else
        strWhere = ""
    End If       
    
    Call SubMakeSQLStatements("MR",strWhere,"X","X")                                '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF                           
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MNU_ID")  )          
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(GetCodeName("1",lgObjRs("MNU_ID")))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MNU_ID")   )         
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("USR_ID"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(GetCodeName("2",lgObjRs("USR_ID"))   )          
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INTERNAL_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(GetCodeName("3",lgObjRs("INTERNAL_CD")))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("AUTH_YN"))
            
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
    Call SubCloseRs(lgObjRs)
End Sub
'============================================================================================================
' Name : GetCodeName(iVal)
' Desc : 
'============================================================================================================
Function GetCodeName(iCase,iVal)
Dim IntRetnm
Dim iWhere
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Select Case ICase
        Case "1"              '메뉴명 
             Call CommonQueryRs(" MNU_NM "," Z_LANG_CO_MAST_MNU "," MNU_ID =  " & FilterVar(iVal, "''", "S") & " AND LANG_CD =  " & FilterVar(gLang , "''", "S") & ""  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)             
             IntRetnm = Trim(Replace(lgF0,Chr(11),""))
        Case "2"              'User명 
            Call CommonQueryRs(" USR_NM "," Z_USR_MAST_REC "," USR_ID =  " & FilterVar(iVal, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            IntRetnm = Trim(Replace(lgF0,Chr(11),""))
        Case "3"              '부서명 
            iWhere = " INTERNAL_CD =  " & FilterVar(iVal, "''", "S") & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"
            Call CommonQueryRs(" Top 1 DEPT_NM "," B_ACCT_DEPT ",  iWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            IntRetnm = Trim(Replace(lgF0,Chr(11),""))
    End Select 
    
    If intRetnm = false Then
        GetCodeName = ""
    Else
        GetCodeName = IntRetnm
    End If
    
End Function
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
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
 
    Dim strCon
    lgStrSQL = "INSERT INTO HZA010T("
    lgStrSQL = lgStrSQL & " MNU_ID,USR_ID,INTERNAL_CD,AUTH_YN,ISRT_EMP_NO,ISRT_DT,UPDT_EMP_NO,UPDT_DT"
    lgStrSQL = lgStrSQL & ")"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S")                         & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S")                         & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(7), "''", "S")                         & ","    
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(8), "''", "S")                         & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                                       & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")							 & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                                       & "," 
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
     
    lgStrSQL = "UPDATE HZA010T"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " MNU_ID= "       & FilterVar(arrColVal(5), "''", "S")                & ","
    lgStrSQL = lgStrSQL & " USR_ID = "      & FilterVar(arrColVal(7), "''", "S")                & ","
    lgStrSQL = lgStrSQL & " INTERNAL_CD = " & FilterVar(arrColVal(8), "''", "S")                & ","
    lgStrSQL = lgStrSQL & " AUTH_YN = "     & FilterVar(arrColVal(9), "''", "S")                & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")                              & "," 
    lgStrSQL = lgStrSQL & " UPDT_DT = "     & FilterVar(GetSvrDateTime, "''", "S")    
    lgStrSQL = lgStrSQL & " WHERE"    
    lgStrSQL = lgStrSQL & " Mnu_ID = "      & FilterVar(UCase(arrColVal(6)), "''", "S")    
    lgStrSQL = lgStrSQL & " AND USR_ID = " & FilterVar(UCase(arrColVal(7)), "''", "S")        

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
  
    lgStrSQL = "DELETE  HZA010T"
    lgStrSQL = lgStrSQL & " WHERE"    
    lgStrSQL = lgStrSQL & " Mnu_ID   = "        & FilterVar(arrColVal(5), "''", "S")    
    lgStrSQL = lgStrSQL & " AND USR_ID   = "    & FilterVar(arrColVal(6), "''", "S")    

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
                       lgStrSQL = "SELECT TOP " & iSelCount  & " MNU_ID,"
                       lgStrSQL = lgStrSQL & " USR_ID,INTERNAL_CD,AUTH_YN"
                       lgStrSQL = lgStrSQL & " FROM HZA010T"
                       lgStrSQL = lgStrSQL & pCode
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
                .ggoSpread.SSShowData "<%=lgstrData%>"                               '☜ : Display data                
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .DBQueryOk        
             End with
          End If
       Case "<%=UID_M0002%>"                                                                '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If
       Case "<%=UID_M0002%>"                                                        '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else
          End If
    End Select
</Script>
