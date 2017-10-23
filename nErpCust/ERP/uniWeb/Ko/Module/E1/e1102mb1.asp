<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
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

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")
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
    Dim iKey1,iRet,ReturnDt,strWhere

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    iRet = SubEmpBase1(lgKeyStream(0),lgKeyStream(1),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
    If iRet = True Then
%>
        <Script Language=vbscript>
            With parent.parent
                .txtEmp_no2.Value = "<%=ConvSPChars(emp_no)%>"
                .txtName2.Value = "<%=ConvSPChars(Name)%>"
            End With
            
            With parent.frm1
                .txtEmp_no.Value = "<%=ConvSPChars(emp_no)%>"
                .txtName.Value = "<%=ConvSPChars(Name)%>"
                .txtDept_nm.value = "<%=ConvSPChars(DEPT_NM)%>"    
                .txtroll_pstn.value = "<%=ConvSPChars(roll_pstn)%>"
            End With          
        </Script>       
<%
    Else
		strWhere = " emp_no=" & FilterVar(lgKeyStream(0), "''", "S")
		strWhere = strWhere & " AND retire_dt is null"       
			
		Call CommonQueryRs(" internal_cd "," HAA010T ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
		if gProAuth = 0 then
			if lgF0="X" or lgF0="" then
	%>
	        <Script Language=vbscript>
	            With parent.parent
	                .txtEmp_no2.Value = "<%=lgKeyStream(0)%>"
	                .txtName2.Value = ""
	   
	            End With
	            With parent.frm1
	                .txtEmp_no.Value = ""
	                .txtName.Value = ""
	                .txtDept_nm.value = ""    
	                .txtroll_pstn.value = ""
	            End With            
	        </Script>       
	<%				
				Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)
				Response.end		
			end if		
		else
			if inStr(ConvSPChars(lgF0),ConvSPChars(lgKeyStream(1)))=0 then
	%>
	        <Script Language=vbscript>
	            With parent.parent
	                .txtEmp_no2.Value = "<%=lgKeyStream(0)%>"
	                .txtName2.Value = ""
	   
	            End With
	            With parent.frm1
	                .txtEmp_no.Value = ""
	                .txtName.Value = ""
	                .txtDept_nm.value = ""    
	                .txtroll_pstn.value = ""
	            End With            
	        </Script>       
	<%		
				if lgF0="X" or lgF0="" then
					Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)
				else 
					Call DisplayMsgBox("800454", vbInformation, "", "", I_MKSCRIPT)
				end if
				Response.end
			end if
		end if
		if  lgPrevNext = "N" then
			Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)
		elseif lgPrevNext = "P" then
			Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)
		end if
		Response.End
    End If

	 iKey1 = FilterVar(lgKeyStream(0), "''", "S")
	 iKey1 = iKey1 & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ""
	 iKey1 = iKey1 & " AND retire_dt is null"    
 
    Call SubMakeSQLStatements("R",iKey1)                                       '¢Ð : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else

'        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)
        lgstrData = ""
        iDx       = 1
        Do While Not lgObjRs.EOF
            emp_no = ConvSPChars(lgObjRs(0))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("gazet_dt"),Null)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gazet_cd"))
            
            ReturnDt = FuncCodeName(2,ConvSPChars(lgObjRs("dept_cd")),UNIConvDateCompanyToDB(lgObjRs("gazet_dt"),NULL))
            
            If ConvSPChars(lgObjRs("gazet_dt")) = ReturnDt Then
				lgstrData = lgstrData & Chr(11) & ""
            Else
				lgstrData = lgstrData & Chr(11) & FuncCodeName(2,ConvSPChars(lgObjRs("dept_cd")),UNIConvDateCompanyToDB(lgObjRs("gazet_dt"),NULL))
            End if
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1")) & "-" & ConvSPChars(lgObjRs("pay_grd2"))
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 

        if  iDx = 1 then
			Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        end if

    End If

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
    Dim strRowBak
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    'iKey1 = FilterVar(lgKeyStream(0),"'%'", "S")
	 iKey1 = FilterVar(lgKeyStream(0), "''", "S")
	 iKey1 = iKey1 & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ""
	 iKey1 = iKey1 & " AND retire_dt is null"    
    Call SubMakeSQLStatements("R",iKey1,"X","like")                                 '¡Ù : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

'        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("eng_name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hanja_name"))
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
               
        Loop 
'        lgstrData = Mid(lgstrData,2)
    End If
    

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet
End Sub    

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

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
                    lgStrSQL = "Select  a.emp_no, gazet_dt, " 
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0029", "''", "S") & ",gazet_cd) as gazet_cd,"
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",a.roll_pstn) as roll_pstn,"
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",a.pay_grd1) as pay_grd1,"
	                lgStrSQL = lgStrSQL & " a.dept_cd, a.pay_grd2"
                    lgStrSQL = lgStrSQL & " From HBA010T a, HAA010T b"
                    lgStrSQL = lgStrSQL & " WHERE a.emp_no *=b.emp_no and a.emp_no = " & pCode
                    lgStrSQL = lgStrSQL & " ORDER BY gazet_dt DESC"
      
                Case "P"
                    lgStrSQL = "Select emp_no, gazet_dt, "    
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0029", "''", "S") & ",gazet_cd) as gazet_cd,"
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) as roll_pstn,"
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",pay_grd1) as pay_grd1,"
	                lgStrSQL = lgStrSQL & " dept_cd,pay_grd2"

                    lgStrSQL = lgStrSQL & " From HBA010T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no = (select top 1 emp_no from haa010t where emp_no < " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & "  and retire_dt is null ORDER BY emp_no DESC)"
                Case "N"
                    lgStrSQL = "Select   emp_no, gazet_dt, "  
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0029", "''", "S") & ",gazet_cd) as gazet_cd,"
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) as roll_pstn,"
	                lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",pay_grd1) as pay_grd1,"
	                lgStrSQL = lgStrSQL & " dept_cd, pay_grd2"
                    lgStrSQL = lgStrSQL & " From HBA010T"
                    lgStrSQL = lgStrSQL & " WHERE emp_no = (select top 1 emp_no from haa010t where emp_no > " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " and retire_dt is null ORDER BY emp_no ASC)"
             End Select
      Case "C"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
      Case "U"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
      Case "D"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
    End Select
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

'Response.Write lgStrSQL
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
					Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
						Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)
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
						Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)
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
             'Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	
