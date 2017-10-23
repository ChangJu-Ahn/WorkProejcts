<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
    call LoadBasisGlobalInf()
    
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
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
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
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
    Dim iCnt, iCnt_holiday
    Dim strDilig_dt
    Dim strWhere, strWhere2

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    If lgKeyStream(3) = "" then
    Else
       strWhere = FilterVar(lgKeyStream(3), "''", "S")
    End if
    
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

    strDilig_dt = " " & FilterVar(uniConvDateCompanyToDB(lgKeyStream(0), gDateFormat), "''", "S") & ""
    
    lgStrSQL =           "SELECT a.emp_no, a.name, d.dept_cd, e.dept_nm, b.wk_type, "
    lgStrSQL = lgStrSQL & "      c.day_time, e.internal_cd internal_cd, f.holi_type, c.holiday_apply "
    lgStrSQL = lgStrSQL & " FROM hdf020t a, hca040t b, hca010t c, hba010t d, b_acct_dept e, hca020t f "
    lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no AND a.emp_no = d.emp_no "
    lgStrSQL = lgStrSQL & " AND (a.retire_dt IS NULL OR a.retire_dt > " & strDilig_dt & ")"
    lgStrSQL = lgStrSQL & " AND c.dilig_cd = " & strWhere
    lgStrSQL = lgStrSQL & " AND a.entr_dt <= " & strDilig_dt
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
    lgStrSQL = lgStrSQL & " AND f.org_cd = dbo.ufn_H_GetCodeName(" & FilterVar("B_COST_CENTER", "''", "S") & ", e.cost_cd, '') "
    lgStrSQL = lgStrSQL & " AND f.wk_type = b.wk_type "
    lgStrSQL = lgStrSQL & " AND f.date = " & strDilig_dt
    lgStrSQL = lgStrSQL & " AND NOT EXISTS (SELECT emp_no " 
    lgStrSQL = lgStrSQL &         " FROM hca060t " 
    lgStrSQL = lgstrSQL &        " WHERE emp_no = a.emp_no " 
    lgStrSQL = lgstrSQL &         "  AND dilig_dt = " & strDilig_dt
    lgStrSQL = lgstrSQL &         "  AND dilig_cd = " & FilterVar(lgKeyStream(3), "''", "S") & ")"
    lgStrSQL = lgStrSQL & " UNION "
    lgStrSQL = lgStrSQL & "SELECT a.emp_no, a.name, a.dept_cd, dbo.ufn_GetDeptName(a.DEPT_CD," & strDilig_dt & ") dept_nm, " 
    lgStrSQL = lgStrSQL & "      b.wk_type, c.day_time, a.internal_cd internal_cd, f.holi_type, c.holiday_apply "
    lgStrSQL = lgStrSQL & " FROM hdf020t a, hca040t b, hca010t c, b_acct_dept e, hca020t f"
    lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no "
    lgStrSQL = lgStrSQL & " AND (a.retire_dt IS NULL OR a.retire_dt > " & strDilig_dt & ")"
    lgStrSQL = lgStrSQL & " AND c.dilig_cd = " & strWhere2
    lgStrSQL = lgStrSQL & " AND a.entr_dt <= " & strDilig_dt
    lgStrSQL = lgstrSQL & " AND a.emp_no NOT IN (SELECT emp_no from hba010t "
    lgStrSQL = lgstrSQL &             " WHERE gazet_dt <= " & strDilig_dt 
    lgStrSQL = lgstrSQL &             "   AND emp_no = a.emp_no " 
    lgStrSQL = lgstrSQL &             "   AND dept_cd is not null) " 
    lgStrSQL = lgstrSQL & " AND b.chang_dt = (SELECT MAX(chang_dt) from hca040t "
    lgStrSQL = lgstrSQL &          " WHERE chang_dt <= " & strDilig_dt 
    lgStrSQL = lgstrSQL &          "   AND emp_no = a.emp_no) " 
    lgStrSQL = lgstrSQL & " AND e.org_change_dt = (SELECT MAX(org_change_dt) from b_acct_dept "
    lgStrSQL = lgstrSQL &               " WHERE dept_cd = a.dept_cd "
    lgStrSQL = lgstrSQL &               "   AND org_change_dt <= " & strDilig_dt & ") " 
    lgStrSQL = lgStrSQL & " AND e.dept_cd = a.dept_cd "
    lgStrSQL = lgStrSQL & " AND f.org_cd = dbo.ufn_H_GetCodeName(" & FilterVar("B_COST_CENTER", "''", "S") & ", e.cost_cd, '') "
    lgStrSQL = lgStrSQL & " AND f.wk_type = b.wk_type "
    lgStrSQL = lgStrSQL & " AND f.date = " & strDilig_dt
    lgStrSQL = lgStrSQL & " AND NOT EXISTS (SELECT emp_no " 
    lgStrSQL = lgStrSQL &         " FROM hca060t " 
    lgStrSQL = lgstrSQL &        " WHERE emp_no = a.emp_no " 
    lgStrSQL = lgstrSQL &         "  AND dilig_dt = " & strDilig_dt
    lgStrSQL = lgstrSQL &         "  AND dilig_cd = " & FilterVar(lgKeyStream(3), "''", "S") & ")"
    lgStrSQL = lgstrSQL & " ORDER BY internal_cd,a.emp_no "

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("800506", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        lgstrData = ""
        iCnt = 0
        iCnt_holiday = 0
        
        Do While Not lgObjRs.EOF
            If Trim(lgObjRs("HOLIDAY_APPLY")) = "N" AND Trim(lgObjRs("HOLI_TYPE")) = "H" Then
                iCnt_holiday = iCnt_holiday + 1
            Else
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
                lgstrData = lgstrData & Chr(11) & ""
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
                lgstrData = lgstrData & Chr(11) & ""
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_TYPE"))
                lgstrData = lgstrData & Chr(11) & ""
                lgstrData = lgstrData & Chr(11) & lgKeyStream(0)
                lgstrData = lgstrData & Chr(11) & lgKeyStream(3)
                lgstrData = lgstrData & Chr(11) & lgKeyStream(4)
                lgstrData = lgstrData & Chr(11) & ""
                lgstrData = lgstrData & Chr(11) & lgKeyStream(5)
                lgstrData = lgstrData & Chr(11) & lgKeyStream(6)
                lgstrData = lgstrData & Chr(11) & lgKeyStream(7)
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_TIME"))

                lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                lgstrData = lgstrData & Chr(11) & Chr(12)

                iCnt = iCnt + 1
            End If

		    lgObjRs.MoveNext
        Loop 

		If iCnt = 0 Then
		    If iCnt_holiday > 0 Then
    		    Call DisplayMsgBox("800505", vbInformation, lgKeyStream(9), "", I_MKSCRIPT)      '¢Ð : No data is found. 
	        Else	    
    		    Call DisplayMsgBox("800065", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
    		End If        

		    Call SetErrorStatus()
		End If        
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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBAutoQueryOk
	         End with
          End If   
    End Select    
    
       
</Script>	
