<%@ LANGUAGE="VBSCRIPT" %>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->

<%

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd_uniSIMS
    Dim emp_no
    Dim name
    Dim dept_nm

    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)

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
    Dim DiligAuth,strDiligAuth,DiligAuths,AuthCheck,AuthChecks
    Dim Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    strDiligAuth = ""

    If  Replace(lgKeyStream(3),chr(11),"")="" Then
        strDiligAuth = " internal_cd LIKE " & FilterVar("%", "''", "S") & ""
    Else
		DiligAuths= lgKeyStream(3)
		AuthChecks = lgKeyStream(4)
		DiligAuth=  mid(DiligAuths,1,instr(1,DiligAuths,chr(12),1)-1)
		DiligAuths = mid(DiligAuths,instr(1,DiligAuths,chr(12),1)+1,len(DiligAuths) - instr(1,DiligAuths,chr(12),1))
		AuthCheck =  mid(AuthChecks,1,instr(1,AuthChecks,chr(12),1)-1)
		AuthChecks = mid(AuthChecks,instr(1,AuthChecks,chr(12),1)+1,len(AuthChecks) - instr(1,AuthChecks,chr(12),1))

		if Trim(DiligAuth) <>"" then
			if Trim(AuthCheck) = "Y" then		
				strDiligAuth =  " internal_cd LIKE " & FilterVar(DiligAuth& "%", "''", "S")
			else
				strDiligAuth =  " internal_cd  LIKE " & FilterVar(DiligAuth , "''", "S")			
			end if

		end if	

		do while instr(1,DiligAuths,chr(12),1)>0
				DiligAuth=  mid(DiligAuths,1,instr(1,DiligAuths,chr(12),1)-1)
				DiligAuths = mid(DiligAuths,instr(1,DiligAuths,chr(12),1)+1,len(DiligAuths) - instr(1,DiligAuths,chr(12),1))
				AuthCheck =  mid(AuthChecks,1,instr(1,AuthChecks,chr(12),1)-1)
				AuthChecks = mid(AuthChecks,instr(1,AuthChecks,chr(12),1)+1,len(AuthChecks) - instr(1,AuthChecks,chr(12),1))
			
				if Trim(DiligAuth) <>"" then
					if Trim(AuthCheck) = "Y" then
						strDiligAuth = strDiligAuth & " or internal_cd  LIKE " & FilterVar(DiligAuth& "%", "''", "S")
					else 
						strDiligAuth = strDiligAuth & " or internal_cd  LIKE " & FilterVar(DiligAuth , "''", "S")			
					end if
				end if			
		Loop
'Response.Write "strDiligAuth:" & strDiligAuth

	end if
	
	if lgKeyStream(0) <>"" then
		iKey1 = iKey1 & " emp_no >= " & FilterVar(lgKeyStream(0), "''", "S")    & " AND "
	end if
	if lgKeyStream(1) <>"" then
		iKey1 = iKey1 & " name LIKE " & FilterVar("%" & lgKeyStream(1) & "%", "''", "S") & " AND "
    end if
    iKey1 = iKey1 & "  retire_dt is null"
    
	if lgKeyStream(2) <>"APPROVAL_CODE" then
		iKey1 = iKey1 & " and (" & strDiligAuth
		iKey1 = iKey1 & "  or emp_no= " & FilterVar(lgKeyStream(5), "''", "S") & ")"
	end if
'Response.Write "**iKey1:" & iKey1	
'Response.End
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
                    lgStrSQL = "Select emp_no,name,dept_nm, res_no, "
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn"
                    lgStrSQL = lgStrSQL & " From HAA010T"
                    lgStrSQL = lgStrSQL & " WHERE " & pCode
                    lgStrSQL = lgStrSQL & " Order by emp_no ASC"
                Case "P"
                    lgStrSQL = "Select emp_no,name,dept_nm, res_no, "
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn"
                    lgStrSQL = lgStrSQL & " From HAA010T"
                    lgStrSQL = lgStrSQL & " WHERE " & pCode
                    lgStrSQL = lgStrSQL & " Order by emp_no ASC"

                Case "N"
                    lgStrSQL = "Select emp_no,name,dept_nm, res_no, "
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn"
                    lgStrSQL = lgStrSQL & " From HAA010T"
                    lgStrSQL = lgStrSQL & " WHERE " & pCode
                    lgStrSQL = lgStrSQL & " Order by emp_no ASC"
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
    End Select    
       
</Script>	
