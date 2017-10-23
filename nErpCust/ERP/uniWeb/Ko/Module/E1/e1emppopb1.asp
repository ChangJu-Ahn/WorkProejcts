<%@ LANGUAGE="VBSCRIPT" %>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd_uniSIMS
    Dim emp_no
    Dim name
    Dim dept_nm
                                                               '��: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        
    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '��: Fetch count at a time for VspdData
'    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '��: Query
             Call SubBizQuery()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
	if gProAuth = 0 then
		iKey1 = " emp_no >= " & FilterVar(lgKeyStream(0), "''", "S") 
		iKey1 = iKey1 & " AND name LIKE " & FilterVar("%" & lgKeyStream(1) & "%", "''", "S")
		iKey1 = iKey1 & " AND retire_dt is null"
	else		
		iKey1 = " emp_no >= " & FilterVar(lgKeyStream(0), "''", "S") 
		iKey1 = iKey1 & " AND name LIKE " & FilterVar("%" & lgKeyStream(1) & "%", "''", "S")
		iKey1 = iKey1 & " AND internal_cd LIKE " & FilterVar(lgKeyStream(2) & "%", "''", "S")
		iKey1 = iKey1 & " AND retire_dt is null"
	end if
    Call SubMakeSQLStatements("R",iKey1)                                       '�� : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
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
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 
    End If

    Call SubCloseRs(lgObjRs)

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
                Case ""
                    lgStrSQL = "Select emp_no,name,dept_nm, res_no, "
	                lgStrSQL = lgStrSQL & " (select b_minor.minor_nm from b_minor where b_minor.minor_cd = roll_pstn and b_minor.major_cd=" & FilterVar("H0002", "''", "S") & ") as roll_pstn"
                    lgStrSQL = lgStrSQL & " From HAA010T"
                    lgStrSQL = lgStrSQL & " WHERE " & pCode
                    lgStrSQL = lgStrSQL & " Order by emp_no ASC"
                Case "P"
                    lgStrSQL = "Select emp_no, "
                    lgStrSQL = lgStrSQL & " trip_dt, trip_cd,"
                    lgStrSQL = lgStrSQL & " (select hca010t.trip_nm from hca010t where hca010t.trip_cd = E11060T.trip_cd) as trip_nm ,"
                    lgStrSQL = lgStrSQL & " trip_cnt, "
                    lgStrSQL = lgStrSQL & " trip_hh, "
                    lgStrSQL = lgStrSQL & " trip_mm, "
                    lgStrSQL = lgStrSQL & " (select haa010t.name from haa010t where haa010t.emp_no = E11060T.app_emp_no) as app_emp_no ,"
                    lgStrSQL = lgStrSQL & " app_yn "
                    lgStrSQL = lgStrSQL & " From E11060T"
                    lgStrSQL = lgStrSQL & " WHERE app_emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & " Order by trip_dt DESC"
                    'lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode
                Case "N"
                    lgStrSQL = "Select emp_no, "
                    lgStrSQL = lgStrSQL & " trip_dt, trip_cd,"
                    lgStrSQL = lgStrSQL & " (select hca010t.trip_nm from hca010t where hca010t.trip_cd = E11060T.trip_cd) as trip_nm ,"
                    lgStrSQL = lgStrSQL & " trip_cnt, "
                    lgStrSQL = lgStrSQL & " trip_hh, "
                    lgStrSQL = lgStrSQL & " trip_mm, "
                    lgStrSQL = lgStrSQL & " (select haa010t.name from haa010t where haa010t.emp_no = E11060T.app_emp_no) as app_emp_no ,"
                    lgStrSQL = lgStrSQL & " app_yn "
                    lgStrSQL = lgStrSQL & " From E11060T"
                    lgStrSQL = lgStrSQL & " WHERE app_emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & " Order by trip_dt DESC"
                    'lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
             End Select
      Case "C"
      Case "U"
      Case "D"
    End Select
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

'Response.Write lgStrSQL
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
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
       Case "UID_M0001"                                                         '�� : Query
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
       Case "UID_M0002"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             'parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	
