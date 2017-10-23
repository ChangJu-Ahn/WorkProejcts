<%@ LANGUAGE="VBSCRIPT" %>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->

<%

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd_uniSIMS
    Dim emp_no
    Dim name
    Dim dept_nm
                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

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
    Dim DiligAuth,strDiligAuth,DiligAuths,AuthCheck,AuthChecks
    Dim Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
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
    Call SubMakeSQLStatements("R",iKey1)                                       '¢Ð : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
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
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 
    End If
    Call SubCloseRs(lgObjRs)

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
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
                Case ""
                    lgStrSQL = "Select emp_no,name,dept_nm, res_no, "
	                lgStrSQL = lgStrSQL & " (select b_minor.minor_nm from b_minor where b_minor.minor_cd = roll_pstn and b_minor.major_cd=" & FilterVar("H0002", "''", "S") & ") as roll_pstn"
                    lgStrSQL = lgStrSQL & " From HAA010T"
                    lgStrSQL = lgStrSQL & " WHERE " & pCode
                    lgStrSQL = lgStrSQL & " Order by emp_no ASC"
                Case "P"
                    lgStrSQL = "Select emp_no,name,dept_nm, res_no, "
	                lgStrSQL = lgStrSQL & " (select b_minor.minor_nm from b_minor where b_minor.minor_cd = roll_pstn and b_minor.major_cd=" & FilterVar("H0002", "''", "S") & ") as roll_pstn"
                    lgStrSQL = lgStrSQL & " From HAA010T"
                    lgStrSQL = lgStrSQL & " WHERE " & pCode
                    lgStrSQL = lgStrSQL & " Order by emp_no ASC"

                Case "N"
                    lgStrSQL = "Select emp_no,name,dept_nm, res_no, "
	                lgStrSQL = lgStrSQL & " (select b_minor.minor_nm from b_minor where b_minor.minor_cd = roll_pstn and b_minor.major_cd=" & FilterVar("H0002", "''", "S") & ") as roll_pstn"
                    lgStrSQL = lgStrSQL & " From HAA010T"
                    lgStrSQL = lgStrSQL & " WHERE " & pCode
                    lgStrSQL = lgStrSQL & " Order by emp_no ASC"
             End Select
      Case "C"
      Case "U"
      Case "D"
    End Select
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
'Response.Write "lgStrSQL:" & lgStrSQL                    	
'Response.End 

End Sub
                     

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
       Case "UID_M0002"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             'parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	
