<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
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
    Dim lgGetSvrDateTime

    call LoadBasisGlobalInf()
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")
        
    lgGetSvrDateTime = GetSvrDateTime
    
	
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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strDilig_dt
    Dim strWhere, strWhere2
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strDilig_dt = FilterVar(UNIConvDateCompanyToDB(lgKeyStream(0), gDateFormat), "''", "S")

    strWhere = FilterVar(lgKeyStream(3), "''", "S")
   
    If lgKeyStream(2) = "" then
       strWhere = strWhere & " AND b.wk_type LIKE " & FilterVar("%", "''", "S") & ""
    Else
       strWhere = strWhere & " AND b.wk_type LIKE " & FilterVar(lgKeyStream(2), "''", "S")
    End if
  
    strWhere2 = strWhere  

    If lgKeyStream(8) = "" then
        strWhere = strWhere & " AND e.internal_cd LIKE  " & FilterVar(Trim(lgKeyStream(1)) & "%", "''", "S") & " " 
    Else
        strWhere = strWhere & " AND e.internal_cd = " & FilterVar(lgKeyStream(1), "''", "S")
    End if 

    If lgKeyStream(8) = "" then
        strWhere2 = strWhere2 & " AND a.internal_cd LIKE  " & FilterVar(Trim(lgKeyStream(1)) & "%", "''", "S") & " " 
    Else
        strWhere2 = strWhere2 & " AND a.internal_cd = " & FilterVar(lgKeyStream(1), "''", "S")
    End if 

    Call SubMakeSQLStatements("MR",strWhere,strWhere2,strDilig_dt,"<=")                                 '¡Ù : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)    
        iDx       = 1
        lgstrData = ""
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_TYPE"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & lgKeyStream(0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_NM"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DILIG_CNT"),ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DILIG_HH"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DILIG_MM"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_TIME"))
            
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
'-----------
	Dim itxtSpread
	Dim itxtSpreadArr
	Dim itxtSpreadArrCount
	Dim iCUCount
	Dim iDCount
	Dim ii
	
	itxtSpread = ""
	             
	iCUCount = Request.Form("txtCUSpread").Count
	iDCount  = Request.Form("txtDSpread").Count

	itxtSpreadArrCount = -1
	             
	ReDim itxtSpreadArr(iCUCount + iDCount)
	             
	For ii = 1 To iDCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
	Next
	For ii = 1 To iCUCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
	Next
 
   itxtSpread = Join(itxtSpreadArr,"")
 '---------      
	arrRowVal = Split(itxtSpread, gRowSep)                                 '¢Ð: Split Row    data
	
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

    lgStrSQL = "INSERT INTO HCA060T("
    lgStrSQL = lgStrSQL & " EMP_NO," 
    lgStrSQL = lgStrSQL & " DILIG_DT," 
    lgStrSQL = lgStrSQL & " DILIG_CD," 
    lgStrSQL = lgStrSQL & " DILIG_CNT,"
    lgStrSQL = lgStrSQL & " DILIG_HH," 
    lgStrSQL = lgStrSQL & " DILIG_MM,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB((arrColVal(3)),NULL), "''", "S")   & "," 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0)     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
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

    lgStrSQL = "UPDATE  HCA060T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " DILIG_CNT  = "    & UNIConvNum(arrColVal(5),0) & ","
    lgStrSQL = lgStrSQL & " DILIG_HH  = "     & UNIConvNum(arrColVal(6),0) & ","
    lgStrSQL = lgStrSQL & " DILIG_MM  = "     & UNIConvNum(arrColVal(7),0) 
    
    lgStrSQL = lgStrSQL & " WHERE EMP_NO = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND DILIG_DT = " & FilterVar(UNIConvDateCompanyToDB((arrColVal(3)),NULL), "''", "S")
    lgStrSQL = lgStrSQL & " AND DILIG_CD = " & FilterVar(Trim(UCase(arrColVal(4))),"'","S")

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

    lgStrSQL = "DELETE  HCA060T "
    lgStrSQL = lgStrSQL & " WHERE EMP_NO = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND DILIG_DT = " & FilterVar(UNIConvDateCompanyToDB((arrColVal(3)),NULL), "''", "S")
    lgStrSQL = lgStrSQL & " AND DILIG_CD = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pWhere1,pWhere2,pCode1,pComp)
    Dim iSelCount
    Dim strDilig_dt
    
    strDilig_dt = pCode1
    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "SELECT  a.emp_no, a.name, d.dept_cd, e.dept_nm, b.wk_type, c.day_time, " 
                       lgStrSQL = lgStrSQL & "      f.dilig_cd, c.dilig_nm, f.dilig_cnt, f.dilig_hh, f.dilig_mm, e.internal_cd internal_cd "
                       lgStrSQL = lgStrSQL & " FROM hdf020t a, hca040t b, hca010t c, hba010t d, b_acct_dept e, hca060t f "
                       lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no AND a.emp_no = d.emp_no "
                       lgStrSQL = lgStrSQL & " AND (a.retire_dt IS NULL OR a.retire_dt > " & strDilig_dt & ")"
                       lgStrSQL = lgStrSQL & " AND a.entr_dt <= " & strDilig_dt
                       lgStrSQL = lgStrSQL & " AND a.emp_no = f.emp_no AND f.dilig_dt = " & strDilig_dt
                       lgStrSQL = lgStrSQL & " AND f.dilig_cd = " & pWhere1
                       lgStrSQL = lgStrSQL & " AND f.dilig_cd = c.dilig_cd"
                       lgStrSQL = lgstrSQL & " AND b.chang_dt = (SELECT MAX(chang_dt) from hca040t "
                       lgStrSQL = lgstrSQL &          " WHERE chang_dt <= " & strDilig_dt 
                       lgStrSQL = lgstrSQL &          "   AND emp_no = a.emp_no) " 
                       lgStrSQL = lgstrSQL & " AND d.gazet_dt = (SELECT MAX(gazet_dt) from hba010t "
                       lgStrSQL = lgstrSQL &          " WHERE gazet_dt <= " & strDilig_dt 
                       lgStrSQL = lgstrSQL &          "   AND emp_no = a.emp_no " 
                       lgStrSQL = lgstrSQL &          "   AND dept_cd is not null) " 
                       lgStrSQL = lgstrSQL & " AND e.org_change_dt = (SELECT MAX(org_change_dt) from b_acct_dept "
                       lgStrSQL = lgstrSQL &               " WHERE dept_cd = d.dept_cd "
                       lgStrSQL = lgstrSQL &               "   AND org_change_dt <= " & strDilig_dt & ") " 
                       lgStrSQL = lgStrSQL & " AND e.dept_cd = d.dept_cd "
                       lgStrSQL = lgstrSQL & " AND a.emp_no IN (SELECT emp_no from hba010t "
                       lgStrSQL = lgstrSQL &         " WHERE gazet_dt <= " & strDilig_dt 
                       lgStrSQL = lgstrSQL &         "   AND emp_no = a.emp_no " 
                       lgStrSQL = lgstrSQL &         "   AND dept_cd is not null) "
                       lgStrSQL = lgStrSQL & " UNION "
                       lgStrSQL = lgStrSQL & "SELECT a.emp_no, a.name, a.dept_cd, dbo.ufn_GetDeptName(a.DEPT_CD," & strDilig_dt & ") dept_nm, " 
                       lgStrSQL = lgStrSQL & "      b.wk_type, c.day_time, f.dilig_cd, c.dilig_nm, f.dilig_cnt, f.dilig_hh, f.dilig_mm, a.internal_cd   internal_cd "
                       lgStrSQL = lgStrSQL & " FROM hdf020t a, hca040t b, hca010t c, hca060t f"
                       lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no "
                       lgStrSQL = lgStrSQL & " AND (a.retire_dt IS NULL OR a.retire_dt > " & strDilig_dt & ")"
                       lgStrSQL = lgStrSQL & " AND a.entr_dt <= " & strDilig_dt
                       lgStrSQL = lgStrSQL & " AND a.emp_no = f.emp_no AND f.dilig_dt = " & strDilig_dt
                       lgStrSQL = lgStrSQL & " AND f.dilig_cd = " & pWhere2
                       lgStrSQL = lgStrSQL & " AND f.dilig_cd = c.dilig_cd"
                       lgStrSQL = lgstrSQL & " AND a.emp_no NOT IN (SELECT emp_no from hba010t "
                       lgStrSQL = lgstrSQL &             " WHERE gazet_dt <= " & strDilig_dt 
                       lgStrSQL = lgstrSQL &             "   AND emp_no = a.emp_no " 
                       lgStrSQL = lgstrSQL &             "   AND dept_cd is not null) " 
                       lgStrSQL = lgstrSQL & " AND b.chang_dt = (SELECT MAX(chang_dt) from hca040t "
                       lgStrSQL = lgstrSQL &          " WHERE chang_dt <= " & strDilig_dt 
                       lgStrSQL = lgstrSQL &          "   AND emp_no = a.emp_no) " 
                       lgStrSQL = lgstrSQL & " ORDER BY internal_cd,a.emp_no "
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
