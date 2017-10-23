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

    call LoadBasisGlobalInf()
	Call loadInfTB19029B( "Q", "H","NOCOOKIE","MB")
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(UNIConvDateCompanyToDB((lgKeyStream(0)),NULL),"NULL","S")
   
    If iKey1 = "" then
    Else
       strWhere = "  sum(CASE WHEN b.retire_dt =  " & iKey1 & " and t1.dilig_cd is null and b.emp_no = t2.emp_no THEN 1  when t1.dilig_cd is not null then null ELSE 0 END), "
       strWhere = strWhere & " sum(CASE WHEN b.entr_dt = "   & iKey1 & " and t1.dilig_cd is null and b.emp_no = t2.emp_no  THEN 1 when t1.dilig_cd is not null then null  else 0 END), t1.dilig_cd  , t1.dilig_nm  , t1.cnt   "
       strWhere = strWhere & "  FROM b_acct_dept a, haa010t b, "
       strWhere = strWhere & " (SELECT d.emp_no, d.dilig_cd, c.dilig_nm, SUM(d.dilig_cnt) cnt  "
       strWhere = strWhere & "  FROM hca010t c, hca060t d WHERE c.dilig_cd = d.dilig_cd AND d.dilig_dt = " & iKey1
       strWhere = strWhere & " AND c.dilig_type = " & FilterVar("1", "''", "S") & "  GROUP BY d.emp_no, d.dilig_cd, c.dilig_nm) AS t1,  "
       strWhere = strWhere & "  (select emp_no, dept_cd from haa010t where (retire_dt >= " & iKey1
       strWhere = strWhere & "  OR retire_dt IS NULL) and entr_dt <=  " & iKey1 & " ) AS t2 "
       strWhere = strWhere & " WHERE b.emp_no *= t1.emp_no  AND (b.retire_dt >= " & iKey1 & " OR b.retire_dt IS NULL)  "
       strWhere = strWhere & " AND b.entr_dt <= " & iKey1 & " AND b.dept_cd = a.dept_cd  and b.dept_cd = t2.dept_cd  and b.internal_cd LIKE  " & FilterVar(Trim(lgKeyStream(1)) & "%", "''", "S") & " " 
       strWhere = strWhere & " AND a.org_change_dt = (SELECT MAX(org_change_dt) FROM b_acct_dept WHERE org_change_dt <= GETDATE())  "
       strWhere = strWhere & " GROUP BY a.dept_cd, a.dept_nm, t1.dilig_cd, t1.dilig_nm, t1.cnt ORDER BY a.dept_cd, t1.dilig_cd  "
    End if

    Call SubMakeSQLStatements("MR",strWhere,"X","")                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

     '합계는 한번에 모두 select 한다...
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(2), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(3), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(4), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(5))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(7), ggAmtOfMoney.DecPoint,0)
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
                       lgStrSQL = "SELECT TOP " & iSelCount  & "a.dept_cd  ,  a.dept_nm  ,  case when t1.dilig_cd is null then COUNT(distinct t2.emp_no)  else count(distinct b.emp_no) end  ,  " & pCode
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
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
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

