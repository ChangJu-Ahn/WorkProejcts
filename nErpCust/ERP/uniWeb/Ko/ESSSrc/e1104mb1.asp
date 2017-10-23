<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->

<%

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear 
                                                                           '¢Ð: Clear Error status
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
    Dim iKey1,iRet,strWhere

    'On Error Resume Next                                                             '¢Ð: Protect system from crashing
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

	iKey1 = FilterVar(lgKeyStream(0), "''", "S") & " AND retire_dt is null"
	If gProAuth <> 0 Then
		iKey1 = iKey1 & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S")
    End If

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
            emp_no = ConvSPChars(lgObjRs("emp_no"))

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("rel_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("family_nm"))
            if LEN(ConvSPChars(lgObjRs("res_no"))) > 0 then
                lgstrData = lgstrData & Chr(11) & Mid(ConvSPChars(lgObjRs("res_no")), 1, 6) & "-" & Mid(ConvSPChars(lgObjRs("res_no")), 7, 7)
            else
                lgstrData = lgstrData & Chr(11) & ""
            end if
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sch_ship"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("occup_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("comp_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("roll_pstn"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("supp_nm"))
            If ConvSPChars(lgObjRs("reside_type")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "Y"
            Else
                lgstrData = lgstrData & Chr(11) & "N"
            End If

            If ConvSPChars(lgObjRs("mdcl_insur")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "Y"
            Else
                lgstrData = lgstrData & Chr(11) & "N"
            End If

            If ConvSPChars(lgObjRs("allow_type")) = "Y" then
                lgstrData = lgstrData & Chr(11) & "Y"
            Else
                lgstrData = lgstrData & Chr(11) & "N"
            End If

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
					lgStrSQL = "SELECT emp_no, family_nm, res_no, occup_nm, comp_nm,"
					lgStrSQL = lgStrSQL & " roll_pstn, dbo.ufn_GetCodeName('H0024', supp_cd) as supp_nm, reside_type, mdcl_insur, allow_type,"
					lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('H0023', rel_cd) as rel_cd, "
					lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('H0007', sch_ship) as sch_ship "
					lgStrSQL = lgStrSQL & " FROM  HAA020T"
                    lgStrSQL = lgStrSQL & " WHERE emp_no = (Select emp_no From haa010t Where emp_no = " & pCode & ")"
					lgStrSQL = lgStrSQL & " ORDER BY rel_cd, res_no ASC"                
				Case "P"
					lgStrSQL = "SELECT emp_no, family_nm, res_no, occup_nm, comp_nm,"
					lgStrSQL = lgStrSQL & " roll_pstn, dbo.ufn_GetCodeName('H0024', supp_cd) as supp_nm, reside_type, mdcl_insur, allow_type,"
					lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('H0023', rel_cd) as rel_cd, "
					lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('H0007', sch_ship) as sch_ship "
					lgStrSQL = lgStrSQL & " FROM  HAA020T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no = (select top 1 emp_no from haa010t where emp_no < " & pCode
                    lgStrSQL = lgStrSQL &				  " ORDER BY emp_no DESC)"
					lgStrSQL = lgStrSQL & " ORDER BY rel_cd, res_no ASC"                
                Case "N"
					lgStrSQL = "SELECT emp_no, family_nm, res_no, occup_nm, comp_nm,"
					lgStrSQL = lgStrSQL & " roll_pstn, dbo.ufn_GetCodeName('H0024', supp_cd) as supp_nm, reside_type, mdcl_insur, allow_type,"
					lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('H0023', rel_cd) as rel_cd, "
					lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('H0007', sch_ship) as sch_ship "
					lgStrSQL = lgStrSQL & " FROM  HAA020T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no = (select top 1 emp_no from haa010t where emp_no > " & pCode
                    lgStrSQL = lgStrSQL &				   " ORDER BY emp_no ASC)"
					lgStrSQL = lgStrSQL & " ORDER BY rel_cd, res_no ASC"                
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
    End Select    
       
</Script>	
