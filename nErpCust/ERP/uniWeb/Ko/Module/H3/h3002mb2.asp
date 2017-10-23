<% Option Explicit%>
<%Response.Buffer = True%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%	
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")  

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    
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
    Dim strWhere
    Dim strGazet_dt
    Dim strDept_cd
    Dim strDept_nm
    Dim strgazet_dept_cd
    Dim strgazet_dept_nm
    Dim strGazet_cd
    Dim strGazet_nm

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1         = FilterVar(lgKeyStream(0), "''", "S")
    strGazet_dt = FilterVar(UNIConvDate(lgKeyStream(1)), "''", "S")
'   ------------일괄생성데이터의 Default값Start
    strDept_cd      = lgKeyStream(2)
    strDept_nm      = lgKeyStream(3)
    strgazet_dept_cd = lgKeyStream(4)
    strgazet_dept_nm = lgKeyStream(5)
    strGazet_cd      = lgKeyStream(6)
    strGazet_nm      = lgKeyStream(7)
'   ------------일괄생성데이터의 Default값End
    strWhere = iKey1
    strWhere = strWhere & " And  entr_dt    <= " & strGazet_dt
    strWhere = strWhere & " And (retire_resn IS NULL OR retire_resn = '') "
    strWhere = strWhere & " And (retire_dt IS NULL OR retire_dt >= " & strGazet_dt & ")"
    strWhere = strWhere & " And emp_no NOT IN (select emp_no from hba010t where hba010t.gazet_dt = " & strGazet_dt
    strWhere = strWhere & " And hba010t.gazet_cd =  " & FilterVar(strGazet_cd , "''", "S") & ")"

    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                              '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("800060", vbInformation, "", "", I_MKSCRIPT)            '☜ : 발령할 사원이 없습니다. 
        Call SetErrorStatus()
    Else

        lgstrData = ""
        iDx       = 1

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd2"))
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dept_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(strDept_nm)

            lgstrData = lgstrData & Chr(11) & ConvSPChars(strgazet_dept_cd)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(strgazet_dept_nm)
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(strGazet_cd)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(strGazet_nm)

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("func_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("role_cd"))

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("comp_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sect_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wk_area_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("flag"))
            
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    iDx =  iDx + 1		    
		    lgObjRs.MoveNext
        Loop
    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

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
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                    lgStrSQL = "Select       name, "
                    lgStrSQL = lgStrSQL & " emp_no, "
                    lgStrSQL = lgStrSQL & " dept_cd, "
                    lgStrSQL = lgStrSQL & " pay_grd1, "
                    lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", pay_grd1) pay_grd1_nm, "
                    lgStrSQL = lgStrSQL & " pay_grd2, "
                    lgStrSQL = lgStrSQL & " roll_pstn,  "
                    lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", roll_pstn) roll_pstn_nm, "
                    lgStrSQL = lgStrSQL & " func_cd,  "
                    lgStrSQL = lgStrSQL & " role_cd,  "
                    lgStrSQL = lgStrSQL & " entr_dt,  "
                    lgStrSQL = lgStrSQL & " resent_promote_dt, "
                    lgStrSQL = lgStrSQL & " " & FilterVar("_i_", "''", "S") & " flag, "
                    lgStrSQL = lgStrSQL & " dept_cd chng_dept_cd, "
                    lgStrSQL = lgStrSQL & " comp_cd,  "
                    lgStrSQL = lgStrSQL & " sect_cd,  "
                    lgStrSQL = lgStrSQL & " wk_area_cd  "
                    lgStrSQL = lgStrSQL & " From  haa010t "
                    lgStrSQL = lgStrSQL & " Where dept_cd " & pComp & pCode

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
                .ggoSpread.SSShowData "<%=lgstrData%>"  
                .DBQueryOk1        
	          End with
	      Else
	            parent.Frm1.btnCb_autoisrt.disabled = False
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
