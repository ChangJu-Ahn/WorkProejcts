<%@ LANGUAGE=VBSCript %>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<%
	lgSvrDateTime = GetSvrDateTime
	Call HideStatusWnd_uniSIMS

    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '☜: Save,Update
             Call SubBizSave()
        Case "UID_M0003"
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim txtEmp_no,txtDept_cd,txtInternal_cd,strWhere

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	
    txtEmp_no  = FilterVar(lgKeyStream(0), "''", "S")
    txtDept_cd = FilterVar(lgKeyStream(1), "''", "S")
    txtInternal_cd =FilterVar(lgKeyStream(2), "''", "S")
    strWhere = txtEmp_no
    strWhere = strWhere & " and a.dept_cd = " & txtDept_cd
    strWhere = strWhere & " and a.Emp_no = b.Emp_no and a.Dept_cd = c.Dept_cd "

    Call SubMakeSQLStatements("R",strWhere)                                       '☜ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
%>
<Script Language=vbscript>
       With Parent	
            .frm1.txtDept_cd.value = "<%=ConvSPChars(lgObjRs("Dept_cd"))%>"
            .frm1.txtDept_nm.value = "<%=ConvSPChars(lgObjRs("Dept_nm"))%>"
            .frm1.txtemp_no1.value = "<%=ConvSPChars(lgObjRs("emp_no"))%>"
            .frm1.txtname1.value   = "<%=ConvSPChars(lgObjRs("name"))%>"
            .frm1.txtuse_ynv.value = "<%=ConvSPChars(lgObjRs("internal_auth"))%>"
       End With          
</Script>       
<%     
    End If
    
    Call SubCloseRs(lgObjRs)
    
End Sub    


'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                             '☜ : Update
              Call SubBizSaveSingleUpdate()
    End Select

End Sub	

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  E11090T"
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(lgKeyStream(0), "''", "S")                              ' 사번char(10)
    lgStrSQL = lgStrSQL & " and   dept_cd = " & FilterVar(lgKeyStream(1), "''", "S")                              ' 사번char(10)

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

   On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '등록된 사번 
'    lgStrSQL = "emp_no = " & FilterVar(Request("txtemp_no1"),"''","S")
'    Call CommonQueryRs(" count(emp_no) "," E11090T ", lgStrSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'    if  Replace(lgF0, Chr(11), "") = "X" then
'    else
'        if Cint(Replace(lgF0, Chr(11), "")) > 0 then
'         Call DisplayMsgBox("800474", vbInformation, "", "", I_MKSCRIPT)  
'            lgErrorStatus = "YES"
'            exit sub
'        end if
'    end if

    '미등록된 사번 
    lgStrSQL = "emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    Call CommonQueryRs(" count(emp_no) "," HAA010T ", lgStrSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  Replace(lgF0, Chr(11), "") = 0 then
        if Cint(Replace(lgF0, Chr(11), "")) > 0 then
        else
			Call DisplayMsgBox("800006", vbInformation, "", "", I_MKSCRIPT)  
            lgErrorStatus = "YES"
            exit sub
        end if
    end if
    '미등록된 부서 

    lgStrSQL = "dept_cd = " & FilterVar(lgKeyStream(1), "''", "S") 
    lgStrSQL = lgStrSQL & " AND org_change_dt = (select max(org_change_dt) from b_acct_dept where org_change_dt<=getdate())"
    Call CommonQueryRs(" count(dept_cd) "," b_acct_dept ", lgStrSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  Replace(lgF0, Chr(11), "") = 0 then
        if Cint(Replace(lgF0, Chr(11), "")) > 0 then
        else
			Call DisplayMsgBox("124600", vbInformation, "", "", I_MKSCRIPT)  
            lgErrorStatus = "YES"
            exit sub
        end if
    end if

    '이미 등록된 관리 부서 
    lgStrSQL = "emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "  and dept_cd = " & FilterVar(lgKeyStream(1), "''", "S") 
    Call CommonQueryRs(" count(emp_no) "," E11090T ", lgStrSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  Replace(lgF0, Chr(11), "") = 0 then
    else
        if Cint(Replace(lgF0, Chr(11), "")) > 0 then
			Call DisplayMsgBox("221505", vbInformation, "", "", I_MKSCRIPT)  
            lgErrorStatus = "YES"
            exit sub
        end if
    end if
   
	Call CommonQueryRs("b_acct_dept.internal_cd "," b_acct_dept, b_company "," b_company.cur_org_change_id = b_acct_dept.org_change_id and b_acct_dept.dept_cd= " & FilterVar(lgKeyStream(1), "''", "S"),  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    

    lgStrSQL = "INSERT INTO E11090T("
    lgStrSQL = lgStrSQL & " emp_no, "
    lgStrSQL = lgStrSQL & " dept_cd, "
    lgStrSQL = lgStrSQL & " internal_cd, "
    lgStrSQL = lgStrSQL & " insrt_user_id, "
    lgStrSQL = lgStrSQL & " insrt_dt, "
    lgStrSQL = lgStrSQL & " updt_user_id,  "
    lgStrSQL = lgStrSQL & " updt_dt,internal_auth ) "
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(replace(lgF0,chr(11),""), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(5), "''", "S")            & ","
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(5), "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(4), "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  E11090T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " Dept_cd = " & FilterVar(lgKeyStream(1), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " updt_user_id      = " & FilterVar(lgKeyStream(5), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " updt_dt          = " & FilterVar(lgSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & " internal_auth          = " & FilterVar(lgKeyStream(4), "''", "S")    
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " and Dept_cd = " & FilterVar(lgKeyStream(3), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
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
                    lgStrSQL = "Select Top 1 a.emp_no,b.name,a.dept_cd,c.dept_nm,a.internal_cd,internal_auth" 
                    lgStrSQL = lgStrSQL & " From  e11090t a, haa010t b, b_acct_dept c "
                    lgStrSQL = lgStrSQL & " WHERE a.emp_no = " & pCode 	
                Case "P"
                    lgStrSQL = "Select TOP 1 uid, emp_no, password, pro_auth, dept_auth, " 
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(emp_no)  as name,internal_auth " 
                    lgStrSQL = lgStrSQL & " From  E11002T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 	
                    lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
                Case "N"
                    lgStrSQL = "Select TOP 1 uid, emp_no, password, pro_auth, dept_auth, " 
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(emp_no)  as name ,internal_auth" 
                    lgStrSQL = lgStrSQL & " From  E11002T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
                    lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC "
             End Select
      Case "C"
      Case "U"
      Case "D"
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
					Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)  
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
					Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)  
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
       Case "UID_M0001"                                                         '☜ : Query
         If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                 .DBQueryOk        
	          End with
	      Else
              With Parent
                 .DBQueryFail
	          End with
          End If   
       Case "UID_M0002"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             'parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "UID_M0003"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	

