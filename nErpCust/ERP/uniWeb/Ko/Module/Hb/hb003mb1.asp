<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    
    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    lgSvrDateTime = GetSvrDateTime

    Call HideStatusWnd                                                               '��: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1, iKey2
    Dim strRoll_pstn
    Dim strPay_grd1
    dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	dim strNat_cd
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iKey1 = FilterVar(lgKeyStream(1), "''", "S")
    iKey2 = FilterVar(lgKeyStream(2), "''", "S")

    Call SubMakeSQLStatements("MR",iKey1,iKey2,C_EQ)                                 '�� : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & "" 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & "" 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("duty_day"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_MONEY"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("prov_tot_amt"))
			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sub_tot_amt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("income_tax"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("res_tax"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("real_prov_amt"))
            
            lgstrData = lgstrData & Chr(11) & C_SHEETMAXROWS_D + iDx
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
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet
     
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    'On Error Resume Next                                                             '��: Protect system from crashing

    Err.Clear                                                                        '��: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	
    For iDx = 1 To UBound(arrRowVal, 1)
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '��: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
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
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "INSERT INTO HDF071T( PAY_YYMM, EMP_NO, DUTY_DAY, PROV_TOT_AMT, SUB_TOT_AMT, INCOME_TAX, RES_TAX, REAL_PROV_AMT," 
    lgStrSQL = lgStrSQL & " Isrt_emp_no, Isrt_dt, Updt_emp_no, Updt_dt      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(lgKeyStream(0)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(3), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(4), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(5), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(6), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(7), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(8), "0"), "0", "N") & ","
    
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    lgStrSQL = "UPDATE  HDF071T"
    lgStrSQL = lgStrSQL & " SET " 
    
    lgStrSQL = lgStrSQL & " DUTY_DAY = " & FilterVar(UNICdbl(arrColVal(3), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & " PROV_TOT_AMT = " & FilterVar(UNICdbl(arrColVal(4), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & " SUB_TOT_AMT = " & FilterVar(UNICdbl(arrColVal(5), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & " INCOME_TAX = " & FilterVar(UNICdbl(arrColVal(6), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & " RES_TAX = " & FilterVar(UNICdbl(arrColVal(7), "0"), "0", "N") & ","
    lgStrSQL = lgStrSQL & " REAL_PROV_AMT = " & FilterVar(UNICdbl(arrColVal(8), "0"), "0", "N") & ","

    lgStrSQL = lgStrSQL & " Updt_emp_no = " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " Updt_dt = " & FilterVar(lgSvrDateTime, "''", "S")  

    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & "         emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "		AND PAY_YYMM = " & FilterVar(UCase(lgKeyStream(0)), "''", "S")     & " "

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "DELETE  HDF071T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "         emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "		AND PAY_YYMM = " & FilterVar(UCase(lgKeyStream(0)), "''", "S")     & " "

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
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
                       lgStrSQL = lgStrSQL & " A.EMP_NO, B.EMP_NM, b.dept_cd, c.dept_nm, A.DUTY_DAY, B.DAY_MONEY, A.PROV_TOT_AMT, A.SUB_TOT_AMT, A.InCOME_TAX, A.RES_TAX, A.real_prov_amt"
                       lgStrSQL = lgStrSQL & " FROM  HDF071T A"
                       lgStrSQL = lgStrSQL & "	LEFT OUTER JOIN  HAA011T B ON A.EMP_NO = B.EMP_NO "
                       lgStrSQL = lgStrSQL & "	left outer join  B_ACCT_DEPT c on b.dept_cd = c.dept_cd and c.org_change_dt = ( select max(org_change_dt) from B_ACCT_DEPT where org_change_dt <= case when b.RETIRE_DT is not null then b.RETIRE_DT else b.ENTR_DT end)"
                       
                       lgStrSQL = lgStrSQL & " WHERE A.PAY_YYMM = " & FilterVar(UCase(lgKeyStream(0)), "''", "S") 
                       
                       If Trim(pCode) <> "''" Then
							lgStrSQL = lgStrSQL & " AND A.emp_no = " & pCode 
					   End If
                       If Trim(pCode1) <> "''" Then
							lgStrSQL = lgStrSQL & " AND B.dept_cd = " & pCode1 
					   End If
                       lgStrSQL = lgStrSQL & " ORDER BY A.emp_no"
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
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
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>
