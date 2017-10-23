<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%
    Dim lgSvrDateTime
    
	lgSvrDateTime = GetSvrDateTime
    Call HideStatusWnd_uniSIMS
    Dim emp_no
    Dim name
    Dim dept_nm
                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        
    lgIntFlgMode      = CInt(Request("txtFlgMode"))                                       '¢Ð: Read Operayion Mode (CREATE, UPDATE)
   
    'Multi SpreadSheet

'    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '¢Ð: Fetch count at a time for VspdData
'    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '¢Ð: Query
             Call SubBizQuery()
        Case "UID_M0002"  
			 Call SubBizSave()  
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1,iRet,strWhere

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
	                '.txtEmp_no2.Value = "<%=lgKeyStream(0)%>"
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

	If 	lgKeyStream(1) = "INSERT" Then
	Else

	    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
	    iKey1 = iKey1 & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ""
	    iKey1 = iKey1 & " AND retire_dt is null"    

	    Call SubMakeSQLStatements("R",iKey1,"X","=")                                 '¡Ù : Make sql statements

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
	             
	            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_cd_nm"))
	            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("edu_start_dt"),Null)
	            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("edu_end_dt"),Null)
	            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_office_nm"))
	            lgstrData = lgstrData & Chr(11) & FuncCodeName(3, "", ConvSPChars(lgObjRs("edu_nat")))
	            'lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_nat_nm"))
	            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_cont"))
	            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("edu_type_nm"))
	            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("edu_score"), 2, 0)
	            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("end_dt"),Null)
	            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("edu_fee"), gAmtDecPoint, 0)
	            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("fee_type"))
	            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("repay_amt"), gAmtDecPoint, 0)
	            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("report_type"))
	            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("add_point"), 2, 0)
		
'------ Developer Coding part (End   ) ------------------------------------------------------------------
	            lgstrData = lgstrData & Chr(11) & Chr(12)
			    lgObjRs.MoveNext
	            iDx =  iDx + 1
	        Loop 
	    End If
	    Call SubCloseRs(lgObjRs)
	End if

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
  
    Select Case lgIntFlgMode
        Case  OPMD_CMODE													  '¢Ð : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                            '¢Ð : Update
              Call SubBizSaveSingleUpdate()
    End Select
    
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
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
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
End Sub
'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
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
' Name : SubBizSaveSingleCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
  
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "INSERT INTO HBA030T("
    lgStrSQL = lgStrSQL & "     emp_no,"
    lgStrSQL = lgStrSQL & "     edu_cd,"
    lgStrSQL = lgStrSQL & "     edu_start_dt,"
    lgStrSQL = lgStrSQL & "     edu_end_dt,"
    lgStrSQL = lgStrSQL & "     edu_office,"
    lgStrSQL = lgStrSQL & "     edu_nat,"
    lgStrSQL = lgStrSQL & "     edu_cont,"
    lgStrSQL = lgStrSQL & "     edu_type,"
    lgStrSQL = lgStrSQL & "     edu_score,"
    lgStrSQL = lgStrSQL & "     end_dt,"
    lgStrSQL = lgStrSQL & "     edu_fee,"
    lgStrSQL = lgStrSQL & "     fee_type,"
    lgStrSQL = lgStrSQL & "     repay_amt,"
    lgStrSQL = lgStrSQL & "     report_type,"
    lgStrSQL = lgStrSQL & "     add_point,"
    lgStrSQL = lgStrSQL & "     ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & "     ISRT_DT     ," 
    lgStrSQL = lgStrSQL & "     UPDT_EMP_NO ," 
    lgStrSQL = lgStrSQL & "     UPDT_DT      )" 
    lgStrSQL = lgStrSQL & "     VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEmp_no"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEdu_cd"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEdu_start_dt"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEdu_end_dt"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEdu_office"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEdu_nat"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEdu_cont"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Request("cboEdu_type"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtEdu_score"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEnd_dt"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtEdu_fee"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("cboFee_type"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtRepay_amt"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("cboReport_type"), "''", "S") & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtAdd_point"),0) & ","
    lgStrSQL = lgStrSQL & " " & FilterVar("unierp", "''", "S") & ", "'FilterVar(gUsrId,"","S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & " " & FilterVar("unierp", "''", "S") & ", "'FilterVar(gUsrId,"","S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
	
'	Response.Write lgStrSQL
        
End Sub



'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode,pCode1,pComp)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
             
                Case ""
                       lgStrSQL = "Select  a.emp_no, b.name,"
                       lgStrSQL = lgStrSQL & "edu_cd, edu_start_dt, edu_end_dt, edu_nat, edu_office,"
                       lgStrSQL = lgStrSQL & "edu_cont, edu_type, edu_fee, edu_score, add_point, end_dt,"
                       lgStrSQL = lgStrSQL & "fee_type, report_type, repay_amt"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0033", "''", "S") & ",edu_cd) edu_cd_nm"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0037", "''", "S") & ",edu_office) edu_office_nm"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0109", "''", "S") & ",edu_type) edu_type_nm"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_h_GetCodeName(" & FilterVar("b_country", "''", "S") & "," & FilterVar("edu_nat", "''", "S") & ",'') edu_nat_nm"
                       lgStrSQL = lgStrSQL & " FROM  HBA030T a, HAA010T b "
                       lgStrSQL = lgStrSQL & " WHERE a.emp_no = " & pCode
                       lgStrSQL = lgStrSQL & "   AND a.emp_no = b.emp_no "

                Case "P"
                       lgStrSQL = "Select  a.emp_no, b.name,"
                       lgStrSQL = lgStrSQL & "edu_cd, edu_start_dt, edu_end_dt, edu_nat, edu_office,"
                       lgStrSQL = lgStrSQL & "edu_cont, edu_type, edu_fee, edu_score, add_point, end_dt,"
                       lgStrSQL = lgStrSQL & "fee_type, report_type, repay_amt"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0033", "''", "S") & ",edu_cd) edu_cd_nm"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0037", "''", "S") & ",edu_office) edu_office_nm"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0109", "''", "S") & ",edu_type) edu_type_nm"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_h_GetCodeName(" & FilterVar("b_country", "''", "S") & "," & FilterVar("edu_nat", "''", "S") & ",'') edu_nat_nm"
                       lgStrSQL = lgStrSQL & " FROM  HBA030T a, HAA010T b "
                       lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no and a.emp_no =(select top 1 emp_no from haa010t where emp_no < " & FilterVar(lgKeyStream(0), "''", "S")
						lgStrSQL = lgStrSQL & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " ORDER BY emp_no DESC)"

                Case "N"
                       lgStrSQL = "Select  a.emp_no, b.name,"
                       lgStrSQL = lgStrSQL & "edu_cd, edu_start_dt, edu_end_dt, edu_nat, edu_office,"
                       lgStrSQL = lgStrSQL & "edu_cont, edu_type, edu_fee, edu_score, add_point, end_dt,"
                       lgStrSQL = lgStrSQL & "fee_type, report_type, repay_amt"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0033", "''", "S") & ",edu_cd) edu_cd_nm"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0037", "''", "S") & ",edu_office) edu_office_nm"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0109", "''", "S") & ",edu_type) edu_type_nm"
                       lgStrSQL = lgStrSQL & ",dbo.ufn_h_GetCodeName(" & FilterVar("b_country", "''", "S") & "," & FilterVar("edu_nat", "''", "S") & ",'') edu_nat_nm"
                       lgStrSQL = lgStrSQL & " FROM  HBA030T a, HAA010T b "
                       lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no and  a.emp_no = (select top 1 emp_no from haa010t where emp_no > " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " ORDER BY emp_no ASC)"

             End Select
      Case "C"
        lgStrSQL = "Select *   " 
        lgStrSQL = lgStrSQL & " From  HBA030T "
        lgStrSQL = lgStrSQL & " Where emp_no " & pComp & pCode
      Case "U"
        lgStrSQL = "Select *   " 
        lgStrSQL = lgStrSQL & " From  HBA030T "
        lgStrSQL = lgStrSQL & " Where emp_no " & pComp & pCode
      Case "D"
        lgStrSQL = "Select *   " 
        lgStrSQL = lgStrSQL & " From  HBA030T "
        lgStrSQL = lgStrSQL & " Where emp_no " & pComp & pCode
    End Select
'	Response.Write "lgStrSQL    :" & lgStrSQL
	'Response.end
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
        Case "SC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "SD"
        Case "SR"
        Case "SU"
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
