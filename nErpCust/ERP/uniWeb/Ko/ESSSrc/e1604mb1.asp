<%@ LANGUAGE=VBSCript%>

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
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

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
    
    Call SubEmpBase(lgKeyStream(0),lgKeyStream(1),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
    
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

    if emp_no = "" then
        return
    end if 

    iKey1 = FilterVar(emp_no, "''", "S")
    iKey1 = iKey1 & " AND emp_no = (select emp_no from haa010t where emp_no = " & FilterVar(emp_no, "''", "S")
    iKey1 = iKey1 & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ")"

    Call SubMakeSQLStatements("R",iKey1)                                       '¢Ð : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
        
        lgstrData = ""
        iDx       = 1
        Do While Not lgObjRs.EOF
            emp_no = ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("dilig_dt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_cnt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_hh"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_mm"))
            lgstrData = lgstrData & Chr(11) & "" 'lgObjRs("pims")

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
                Case ""
					lgStrSQL = "Select emp_no, "
					lgStrSQL = lgStrSQL & " dilig_dt, dilig_cd,"
					lgStrSQL = lgStrSQL & " dbo.ufn_H_GetCodeName(" & FilterVar("hca010t", "''", "S") & ",dilig_cd,'') as dilig_nm , "
					lgStrSQL = lgStrSQL & " dilig_cnt, "
					lgStrSQL = lgStrSQL & " dilig_hh, "
					lgStrSQL = lgStrSQL & " dilig_mm "
					lgStrSQL = lgStrSQL & " From HCA060T"
					lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(emp_no, "''", "S")
					lgStrSQL = lgStrSQL & "   AND dilig_dt between" & FilterVar(lgKeyStream(2), "''", "S") & " and " & FilterVar(lgKeyStream(3), "''", "S")
					lgStrSQL = lgStrSQL & "   AND dilig_cd LIKE  " & FilterVar(lgKeyStream(4) & "%", "''", "S") & ""
					lgStrSQL = lgStrSQL & " Order by dilig_dt DESC"
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
			        Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
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
			        Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
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
