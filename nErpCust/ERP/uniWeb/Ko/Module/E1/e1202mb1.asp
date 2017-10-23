<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
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

    Dim TotpayAmt
    Dim TotBonusAmt

    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

'    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '��: Fetch count at a time for VspdData
'    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    

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
    Dim iKey1,iRet
    Dim TotpayAmt, TotBonusAmt
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iRet = SubEmpBase1(lgKeyStream(0),lgKeyStream(1),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)

    If iRet = True Then
%>
        <Script Language=vbscript>
            With parent.frm1
                .txtEmp_no.Value = "<%=ConvSPChars(emp_no)%>"
                .txtName.Value = "<%=ConvSPChars(Name)%>"
                .txtDept_nm.value = "<%=ConvSPChars(DEPT_NM)%>"    
                .txtroll_pstn.value = "<%=ConvSPChars(roll_pstn)%>"
            End With          
        </Script>       
<%
    Else
            if  lgPrevNext = "N" then
                Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)
            elseif lgPrevNext = "P" then
                Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)
            end if
            Response.End
    End If
    iKey1 = " PAY_YYMM  LIKE  " & FilterVar(lgKeyStream(2) & "%", "''", "S") & ""
    iKey1 = iKey1 & "   AND A.EMP_NO = " & FilterVar(emp_no, "''", "S")

    Call SubMakeSQLStatements("R",iKey1)                                       '�� : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
'        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)
        TotpayAmt = 0
        TotBonusAmt = 0
        lgstrData = ""
        iDx       = 1
        Do While Not lgObjRs.EOF
          '  emp_no = ConvSPChars(lgObjRs("emp_no"))
            TotpayAmt = TotpayAmt + CDbl(lgObjRs("PAY_TOT_AMT"))
            TotBonusAmt = TotBonusAmt + CDbl(lgObjRs("BONUS_TOT_AMT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_yymm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TOT_AMT"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BONUS_TOT_AMT"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TAX_AMT"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("NON_TOT_TAX"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("INCOME_TAX"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RES_TAX"), ggAmtOfMoney.DecPoint, 0)

            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 
%>
<Script Language=vbscript>
    With parent.frm1
        .txtTotpayAmt.Value = "<%=UNINumClientFormat(TotpayAmt, ggAmtOfMoney.DecPoint, 0)%>"
        .txtTotBonusAmt.Value = "<%=UNINumClientFormat(TotBonusAmt, ggAmtOfMoney.DecPoint, 0)%>"
    End With          
</Script>       
<%
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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim strRowBak
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet
End Sub    

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '��: Protect system from crashing

    Err.Clear                                                                        '��: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
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
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

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
                Case ""
                Case "P"
                Case "N"
             End Select
      Case "C"
      Case "U"
      Case "D"
    End Select

	lgStrSQL = "Select A.emp_no, pay_yymm,PROV_DT, PROV_TYPE, MINOR_NM, PAY_TOT_AMT, BONUS_TOT_AMT, "
	lgStrSQL = lgStrSQL & "       NON_TAX1 + NON_TAX2 + NON_TAX3 + NON_TAX4 + NON_TAX5 NON_TOT_TAX, "
	lgStrSQL = lgStrSQL & "       TAX_AMT, INCOME_TAX, RES_TAX, SAVE_FUND "
    lgStrSQL = lgStrSQL & " From  HDF070T A, B_MINOR B , HDA270T C "
	lgStrSQL = lgStrSQL & " WHERE " & pCode
    lgStrSQL = lgStrSQL & "   AND PROV_TYPE NOT IN (" & FilterVar("P", "''", "S") & "," & FilterVar("Q", "''", "S") & "," & FilterVar("R", "''", "S") & "," & FilterVar("S", "''", "S") & ") "
    lgStrSQL = lgStrSQL & "   AND MAJOR_CD  = " & FilterVar("H0040", "''", "S") & ""
    lgStrSQL = lgStrSQL & "   AND PROV_TYPE = MINOR_CD "
    lgStrSQL = lgStrSQL & "   AND (((A.PROV_TYPE = C.PAY_TYPE AND C.CLOSE_TYPE = " & FilterVar("2", "''", "S") & " AND A.PAY_YYMM <= convert(varchar(6),C.CLOSE_DT,112)) OR (A.PROV_TYPE = " & FilterVar("1", "''", "S") & " AND C.PAY_TYPE = " & FilterVar("!", "''", "S") & "   AND C.CLOSE_TYPE = " & FilterVar("2", "''", "S") & " AND A.PAY_YYMM <= convert(varchar(6),C.CLOSE_DT,112)) or  (A.PROV_TYPE = " & FilterVar("Z", "''", "S") & " AND C.PAY_TYPE = " & FilterVar("@", "''", "S") & "  AND C.CLOSE_TYPE = " & FilterVar("2", "''", "S") & " AND A.PAY_YYMM <= convert(varchar(6),C.CLOSE_DT,112)))"
    lgStrSQL = lgStrSQL & "   OR  ((A.PROV_TYPE = C.PAY_TYPE AND C.CLOSE_TYPE = " & FilterVar("1", "''", "S") & " AND A.PAY_YYMM < convert(varchar(6),C.CLOSE_DT,112)) OR (A.PROV_TYPE = " & FilterVar("1", "''", "S") & " AND C.PAY_TYPE = " & FilterVar("!", "''", "S") & "   AND C.CLOSE_TYPE = " & FilterVar("1", "''", "S") & " AND A.PAY_YYMM < convert(varchar(6),C.CLOSE_DT,112)) or (A.PROV_TYPE = " & FilterVar("Z", "''", "S") & " AND C.PAY_TYPE = " & FilterVar("@", "''", "S") & "  AND C.CLOSE_TYPE = " & FilterVar("1", "''", "S") & " AND A.PAY_YYMM < convert(varchar(6),C.CLOSE_DT,112))))"
    lgStrSQL = lgStrSQL & " ORDER BY PAY_YYMM DESC, PROV_TYPE"
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
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
