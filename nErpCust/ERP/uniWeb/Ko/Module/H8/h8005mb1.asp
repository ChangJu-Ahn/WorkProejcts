<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

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
' Name : SubBizQueryMulti()
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim strWhere

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strWhere = FilterVar(lgKeyStream(0), "''", "S")   
    strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(3),"'%'", "S") & ") >= " & FilterVar(lgKeyStream(1), "''", "S")         
    strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(3),"'%'", "S") & ") <= " & FilterVar(lgKeyStream(2), "''", "S")            
    strWhere = strWhere & " AND b.emp_no = a.emp_no"
    strWhere = strWhere & " AND a.emp_no IN (SELECT emp_no FROM hdf070t WHERE pay_yymm=" & FilterVar(lgKeyStream(0), "''", "S") & " And prov_type=" & FilterVar("P", "''", "S") & " )"
    strWhere = strWhere & " AND c.code_type=" & FilterVar("1", "''", "S") & " "
    strWhere = strWhere & " AND c.allow_cd=a.allow_cd"
    strWhere = strWhere & " AND a.prov_type IN (" & FilterVar("1", "''", "S") & " ," & FilterVar("P", "''", "S") & " ," & FilterVar("B", "''", "S") & " )"
    strWhere = strWhere & " group by b.dept_cd order by b.dept_cd"    
     
	Call SubMakeSQLStatements("MR",strWhere,"X","=")                                 '¡Ù : Make sql statements       

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""               
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()              
        Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
        Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet 
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        iDx       = 1

        Do While Not lgObjRs.EOF
                 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("EMP_CNT_AMT"), 0,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("COMPUTE_0003_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("COMPUTE_0004_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("COMPUTE_0005_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("COMPUTE_0006_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RETRO1_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RETRO2_AMT"), ggAmtOfMoney.DecPoint,0)
                
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
                       lgStrSQL = " SELECT top " & iSelCount & " (select x.dept_nm from  b_acct_dept x,b_company y where x.org_change_id=y.cur_org_change_id and x.dept_cd=b.dept_cd) AS DEPT_CD, "
                       lgStrSQL = lgStrSQL & " COUNT(DISTINCT b.emp_no) AS EMP_CNT_AMT,"
                       lgStrSQL = lgStrSQL & " CONVERT(NUMERIC(18,4), SUM(CASE a.prov_type WHEN " & FilterVar("1", "''", "S") & "  THEN (CASE c.allow_kind WHEN " & FilterVar("1", "''", "S") & " " 
                       lgStrSQL = lgStrSQL & " THEN ISNULL(a.allow,0) ELSE 0 END) ELSE 0 END)) as COMPUTE_0003_AMT,"
                       lgStrSQL = lgStrSQL & " CONVERT(NUMERIC(18,4), SUM(CASE a.prov_type WHEN " & FilterVar("1", "''", "S") & "  THEN (CASE c.allow_kind WHEN " & FilterVar("2", "''", "S") & ""
                       lgStrSQL = lgStrSQL & " THEN ISNULL(a.allow,0) ELSE 0 END) ELSE 0 END)) as COMPUTE_0004_AMT,"
                       lgStrSQL = lgStrSQL & " CONVERT(NUMERIC(18,4), SUM(CASE a.prov_type WHEN " & FilterVar("P", "''", "S") & "  THEN (CASE c.allow_kind WHEN " & FilterVar("1", "''", "S") & " "
                       lgStrSQL = lgStrSQL & " THEN ISNULL(a.allow,0) ELSE 0 END) ELSE 0 END)) as COMPUTE_0005_AMT,"
                       lgStrSQL = lgStrSQL & " CONVERT(NUMERIC(18,4), SUM(CASE a.prov_type WHEN " & FilterVar("P", "''", "S") & "  THEN (CASE c.allow_kind WHEN " & FilterVar("2", "''", "S") & ""
                       lgStrSQL = lgStrSQL & " THEN ISNULL(a.allow,0) ELSE 0 END) ELSE 0 END)) as COMPUTE_0006_AMT,"
                       lgStrSQL = lgStrSQL & " (CONVERT(NUMERIC(18,4), SUM(CASE a.prov_type WHEN " & FilterVar("P", "''", "S") & "  THEN (CASE c.allow_kind WHEN " & FilterVar("1", "''", "S") & " "
                       lgStrSQL = lgStrSQL & " THEN ISNULL(a.allow,0) ELSE 0 END) ELSE 0 END))"
                       lgStrSQL = lgStrSQL & " - CONVERT(NUMERIC(18,4), SUM(CASE a.prov_type WHEN " & FilterVar("1", "''", "S") & "  THEN (CASE c.allow_kind WHEN " & FilterVar("1", "''", "S") & " "
                       lgStrSQL = lgStrSQL & " THEN ISNULL(a.allow,0) ELSE 0 END) ELSE 0 END))) as RETRO1_AMT,"
                       lgStrSQL = lgStrSQL & " (CONVERT(NUMERIC(18,4), SUM(CASE a.prov_type WHEN " & FilterVar("P", "''", "S") & "  THEN (CASE c.allow_kind WHEN " & FilterVar("2", "''", "S") & ""
                       lgStrSQL = lgStrSQL & " THEN ISNULL(a.allow,0) ELSE 0 END) ELSE 0 END))"
                       lgStrSQL = lgStrSQL & " - CONVERT(NUMERIC(18,4), SUM(CASE a.prov_type WHEN " & FilterVar("1", "''", "S") & "  THEN (CASE c.allow_kind WHEN " & FilterVar("2", "''", "S") & ""
                       lgStrSQL = lgStrSQL & " THEN ISNULL(a.allow,0) ELSE 0 END) ELSE 0 END))) as RETRO2_AMT"
                       lgStrSQL = lgStrSQL & " From  hdf040t a, hdf020t b, hda010t c"                       
                       lgStrSQL = lgStrSQL & " WHERE  a.pay_yymm" & pComp & pCode                       
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
	      Else
	      
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
