<%@ LANGUAGE=VBSCript%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    call LoadBasisGlobalInf()
	Call loadInfTB19029B( "Q", "H","NOCOOKIE","MB")
    
    Call HideStatusWnd                                                               '☜: Hide Processing message

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
    Dim iKey1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = "  " & FilterVar(lgKeyStream(0) & "%", "''", "S") & ""
    if  lgKeyStream(6) = "" then    ' 부서코드를 선택하지 않았을 경우 
        iKey1 = iKey1 + " AND B.INTERNAL_CD LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ""
    end if
    iKey1 = iKey1 + " AND A.LANG_CD LIKE  " & FilterVar(lgKeyStream(2) & "%", "''", "S") & ""
    iKey1 = iKey1 + " AND A.LANG_TYPE LIKE  " & FilterVar(lgKeyStream(3) & "%", "''", "S") & ""
    if  lgKeyStream(4) <> "" then
        if  lgKeyStream(5) = "1" then
            iKey1 = iKey1 + " AND (A.VAL_DT <=  " & FilterVar(UNIConvDate(lgKeyStream(4)), "''", "S") & " or A.VAL_DT is null)"
        elseif lgKeyStream(5) = "2" then
            iKey1 = iKey1 + " AND (A.VAL_DT >=  " & FilterVar(UNIConvDate(lgKeyStream(4)), "''", "S") & " or A.VAL_DT is null)"
        End if
    end if

    Call SubMakeSQLStatements("MR",iKey1,"X", "like")                                 '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lang_nm"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("get_dt"),"")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lang_type_nm"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("score"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("grade"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("val_dt"),"")
            
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

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

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
                       lgStrSQL = "SELECT "
                       lgStrSQL = lgStrSQL & "name, "
                       lgStrSQL = lgStrSQL & "b.emp_no, "
                       lgStrSQL = lgStrSQL & "dept_nm, "
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0058", "''", "S") & ",lang_cd) lang_nm, "                       
                       lgStrSQL = lgStrSQL & "get_dt, "                       
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0059", "''", "S") & ",lang_type) lang_type_nm, "                       
                       lgStrSQL = lgStrSQL & "score, "
                       lgStrSQL = lgStrSQL & "grade, "
                       lgStrSQL = lgStrSQL & "val_dt "                       									
                       lgStrSQL = lgStrSQL & " FROM  HBA060T A, HAA010T B "
                       lgStrSQL = lgStrSQL & " WHERE A.EMP_NO = B.EMP_NO "
                       lgStrSQL = lgStrSQL & " AND   A.EMP_NO " & pComp & pCode
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

Function FuncCodeName(intSW, Major, Minor)
    Dim pRs
    Dim fncSQL
    Select Case intSW
        Case 1  ' B_MAJOR
            fncSQL = "SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD =  " & FilterVar(Major , "''", "S") & " AND MINOR_CD =  " & FilterVar(minor , "''", "S") & ""
        Case 2  ' BCB020T : 부서코드 
            fncSQL = "SELECT DEPT_NAME FROM BCB020T WHERE DEPT_CD =  " & FilterVar(minor , "''", "S") & ""
        Case 3  ' B_COUNTRY : 국적 
            fncSQL = "SELECT COUNTRY_NM FROM B_COUNTRY WHERE COUNTRY_CD =  " & FilterVar(minor , "''", "S") & ""
        Case 4  ' B_COMPANY : 회사코드 
            fncSQL = "SELECT CO_NM FROM B_COMPANY WHERE CO_CD =  " & FilterVar(minor , "''", "S") & ""
	End Select


    If 	FncOpenRs("R",lgObjConn,pRs,fncSQL,"X","X") = False Then
'    If 	FncOpenRs("R",pRs,fncSQL,"X","X") = False Then
        FuncCodeName = Minor
    Else
        FuncCodeName = pRs(0)
    End If

'    If CheckSQLError(pObjRs.ActiveConnection) = True Then
'       ObjectContext.SetAbort
'    End If
'    If CheckSYSTEMError(Err) = True Then
'       ObjectContext.SetAbort
'    End If

End Function



%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
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
