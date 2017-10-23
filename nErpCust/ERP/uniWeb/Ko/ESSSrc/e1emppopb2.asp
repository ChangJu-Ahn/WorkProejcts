<%@ LANGUAGE="VBSCRIPT" %>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->

<%

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd_uniSIMS
    Dim emp_no
    Dim name
    Dim dept_nm

    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '¢Ð: Query
             Call SubBizQuery()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iKey1 = " emp_no >= " & FilterVar(lgKeyStream(0), "''", "S") 
    iKey1 = iKey1 & " AND name LIKE " & FilterVar("%" & lgKeyStream(1) & "%", "''", "S")
    iKey1 = iKey1 & " AND retire_dt is null"

    lgPrevNext = lgKeyStream(3)

    Call SubMakeSQLStatements("R",iKey1)                                       '¢Ð : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
        lgstrData = ""
        iDx       = 1
        Do While Not lgObjRs.EOF


            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("res_no"))

            if  isnull(ConvSPChars(lgObjRs("UID"))) OR  ConvSPChars(lgObjRs("UID")) = "" then
                lgstrData = lgstrData & Chr(11) & ""
            else
                lgstrData = lgstrData & Chr(11) & "µî·Ï"
            end if

            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 
        
    End If
    
    Call SubCloseRs(lgObjRs)

End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount

    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
                Case "A"
                    lgStrSQL = "Select HAA010T.emp_no,HAA010T.name,HAA010T.dept_nm, HAA010T.res_no, "
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn,"
                    lgStrSQL = lgStrSQL & " E11002T.UID"
                    lgStrSQL = lgStrSQL & " From HAA010T, E11002T"
                    lgStrSQL = lgStrSQL & " WHERE HAA010T.emp_no >= " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & "   AND HAA010T.name LIKE " & FilterVar("%" & lgKeyStream(1) & "%", "''", "S")
                    lgStrSQL = lgStrSQL & "   AND HAA010T.retire_dt is null"
                    lgStrSQL = lgStrSQL & "   AND HAA010T.emp_no *= E11002T.emp_no"
                    lgStrSQL = lgStrSQL & " Order by HAA010T.emp_no ASC"

                Case "Y"
                    lgStrSQL = "Select HAA010T.emp_no,HAA010T.name,HAA010T.dept_nm, HAA010T.res_no, "
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn,"
                    lgStrSQL = lgStrSQL & " E11002T.UID"
                    lgStrSQL = lgStrSQL & " From HAA010T, E11002T"
                    lgStrSQL = lgStrSQL & " WHERE HAA010T.emp_no >= " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & "   AND HAA010T.name LIKE " & FilterVar("%" & lgKeyStream(1) & "%", "''", "S")
                    lgStrSQL = lgStrSQL & "   AND HAA010T.retire_dt is null"
                    lgStrSQL = lgStrSQL & "   AND HAA010T.emp_no = E11002T.emp_no"
                    lgStrSQL = lgStrSQL & " Order by HAA010T.emp_no ASC"
                    
                Case else
                    lgStrSQL = "Select HAA010T.emp_no,HAA010T.name,HAA010T.dept_nm, HAA010T.res_no, "
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn,"
	                lgStrSQL = lgStrSQL & " '' as UID"
                    lgStrSQL = lgStrSQL & " From HAA010T"
                    lgStrSQL = lgStrSQL & " WHERE HAA010T.emp_no >= " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & "   AND HAA010T.name LIKE " & FilterVar("%" & lgKeyStream(1) & "%", "''", "S")
                    lgStrSQL = lgStrSQL & "   AND HAA010T.retire_dt is null"
                    lgStrSQL = lgStrSQL & "   AND not exists (select UID from E11002T where E11002T.emp_no = HAA010T.emp_no)"
                    lgStrSQL = lgStrSQL & " Order by HAA010T.emp_no ASC"
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
'                      Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
'                      Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "UID_M0001"                                                         '¢Ð : Query
         If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                 .grid1.SSSetData("<%=lgstrData%>")
                 .DBQueryOk        
	          End with
	      Else
              With Parent
                 .DBQueryFail
	          End with
          End If   
    End Select    
       
</Script>	
