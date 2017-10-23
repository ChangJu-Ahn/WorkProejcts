<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Dim lgSvrDateTime
	lgSvrDateTime = GetSvrDateTime
    Call HideStatusWnd_uniSIMS
                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        
    'Multi SpreadSheet

'    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
'    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '☜: Save,Update
        
             Call SubBizSaveSingleUpdate()
'             Call SubBizSaveMulti()
'        Case CStr(UID_M0003)                                                         '☜: Delete
'             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	if gProAuth = 0 then
		iKey1 = FilterVar(lgKeyStream(0), "''", "S")
	else	
		iKey1 = FilterVar(lgKeyStream(0), "''", "S")
		iKey1 = iKey1 & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ""
	end if
	iKey1 = iKey1 & " AND retire_dt is null"		
	 
    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then      
    	If gProAuth = 0 Then
            %>
            <Script Language=vbscript>
                With parent.parent
                    .txtName2.Value = ""
                End With
    			With Parent
    				.FncNew()    
    			End With             
            </Script>       
            <%		
   	
    	Else
		    strWhere = " emp_no=" & FilterVar(lgKeyStream(0), "''", "S")
		    strWhere = strWhere & " AND retire_dt is null"   
		
		    Call CommonQueryRs(" internal_cd "," HAA010T ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 

    		if lgF0="X" or lgF0="" then
    			Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)
            %>
             <Script Language=vbscript>
                With parent.parent
                    .txtName2.Value = ""
                End With
                With parent
    				.FncNew()    
                End With            
            </Script>   
            <%				
    			Response.end		
    		end if

    		if inStr(1,ConvSPChars(lgF0),ConvSPChars(lgKeyStream(1)))=0 then
    	
            %>
             <Script Language=vbscript>
                With parent.parent
                    .txtName2.Value = ""
                End With
                With parent.frm1
    				.FncNew()    
                End With            
            </Script>   
            <%		
        		Call DisplayMsgBox("800454", vbInformation, "", "", I_MKSCRIPT)
    			Response.end		
            end if  
        End If
            
        If lgPrevNext = "" Then
            Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)
            Call SetErrorStatus()
        ElseIf lgPrevNext = "P" Then
            Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)     '☜ : This is the starting data. 
            lgPrevNext = ""
            Call SubBizQuery()
        ElseIf lgPrevNext = "N" Then
            Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)     '☜ : This is the ending data.
            lgPrevNext = ""
            Call SubBizQuery()
        End If
       
    Else
%>
   
<%		
  
		If ConvSPChars(lgObjRs("mil_type"))="" and ConvSPChars(lgObjRs("mil_kind"))="" and ConvSPChars(lgObjRs("mil_start"))="" and ConvSPChars(lgObjRs("mil_end"))="" and ConvSPChars(lgObjRs("mil_grade"))="" and ConvSPChars(lgObjRs("mil_branch"))="" and ConvSPChars(lgObjRs("mil_no"))=""  Then
            Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
            Call SetErrorStatus()
		End if

%>
<Script Language=vbscript>
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	' Set condition area, contents area
	'--------------------------------------------------------------------------------------------------------
       With parent.parent
            .txtEmp_no2.Value = "<%=ConvSPChars(lgObjRs("emp_no"))%>"
            .txtName2.Value = "<%=ConvSPChars(lgObjRs("name"))%>"
       End With   
       With Parent	
            .Frm1.txtEmp_no.Value  = "<%=ConvSPChars(lgObjRs("emp_no"))%>"
            .Frm1.txtName.Value  = "<%=ConvSPChars(lgObjRs("name"))%>"
            .frm1.txtDept_nm.value = "<%=ConvSPChars(lgObjRs("DEPT_NM"))%>"    
            '.frm1.txtroll_pstn.value = "<%=ConvSPChars(FuncCodeName(1, "H0002", lgObjRs("roll_pstn")))%>"
            .frm1.txtroll_pstn.value = "<%=ConvSPChars(lgObjRs("roll_pstn_nm"))%>"
            
            .frm1.txtmil_type.value = "<%=ConvSPChars(lgObjRs("mil_type"))%>"         
            .frm1.txtmil_kind.value = "<%=ConvSPChars(lgObjRs("mil_kind"))%>"
            .frm1.txtmil_start.value = "<%=UniConvDateDbToCompany(lgObjRs("mil_start"),"")%>"
            .frm1.txtmil_end.value = "<%=UniConvDateDbToCompany(lgObjRs("mil_end"),"")%>"
            .frm1.txtmil_grade.value = "<%=ConvSPChars(lgObjRs("mil_grade"))%>"        
            .frm1.txtmil_branch.value = "<%=ConvSPChars(lgObjRs("mil_branch"))%>"
            .frm1.txtmil_no.value = "<%=ConvSPChars(lgObjRs("mil_no"))%>"              


       End With          
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
End Sub    

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  HAA010T"
    lgStrSQL = lgStrSQL & " SET " 
    '병역구분 char(2)
    lgStrSQL = lgStrSQL & " mil_type = " & FilterVar(UCase(Request("txtmil_type")), "''", "S") & ","
    '병역군별 char(2)
    lgStrSQL = lgStrSQL & " mil_kind = " & FilterVar(UCase(Request("txtmil_kind")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " mil_start = " & FilterVar(UNIConvDateCompanyToDB(Request("txtmil_start"),NULL), "''", "S") & ","    ' datetime
    lgStrSQL = lgStrSQL & " mil_end = " & FilterVar(UNIConvDateCompanyToDB(Request("txtmil_end"),NULL), "''", "S") & ","        ' datetime
    '병역등급 char(2)
    lgStrSQL = lgStrSQL & " mil_grade = " & FilterVar(UCase(Request("txtmil_grade")), "''", "S") & ","
    '병역병과 char(2)
    lgStrSQL = lgStrSQL & " mil_branch = " & FilterVar(UCase(Request("txtmil_branch")), "''", "S") & ","

    lgStrSQL = lgStrSQL & " updt_emp_no = " & FilterVar(gUsrId, "''", "S") & ","                                   ' char(10)
    lgStrSQL = lgStrSQL & " updt_dt = " & FilterVar(lgSvrDateTime, "''", "S") & ","                ' datetime
    '군번 char(10)
    lgStrSQL = lgStrSQL & " mil_no = " & FilterVar(Request("txtmil_no"), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no =  " & FilterVar(Request("txtEmp_no"), "''", "S") & ""

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

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  hba010t"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " dept_cd          = " & FilterVar(UCase(arrColVal(5)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " sect_cd          = " & FilterVar(UCase(arrColVal(6)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " role_cd          = " & FilterVar(UCase(arrColVal(7)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " roll_pstn        = " & FilterVar(UCase(arrColVal(8)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " pay_grd1         = " & FilterVar(UCase(arrColVal(9)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " pay_grd2         = " & FilterVar(UCase(arrColVal(10)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " func_cd          = " & FilterVar(UCase(arrColVal(11)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " gazet_resn       = " & FilterVar(UCase(arrColVal(12)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " updt_emp_no      = " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " updt_dt          = " & FilterVar(lgSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " emp_no           = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " And gazet_dt     = " & FilterVar(UNIConvDate(arrColVal(3)),NULL,"S")
    lgStrSQL = lgStrSQL & " And gazet_cd     = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  hba010t"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " emp_no          = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " And gazet_dt    = " & FilterVar(UNIConvDate(arrColVal(3)),NULL,"S")
    lgStrSQL = lgStrSQL & " And gazet_cd     = " & FilterVar(UCase(arrColVal(4)), "''", "S")
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
                      lgStrSQL = "Select mil_type,mil_kind,mil_start,mil_end,mil_grade,mil_branch,mil_no,emp_no,name,DEPT_NM"
                      lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn_nm "
                      lgStrSQL = lgStrSQL & " From  HAA010T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode 	
                      
                Case "P"
                      lgStrSQL = "Select TOP 1 mil_type,mil_kind,mil_start,mil_end,mil_grade,mil_branch,mil_no,emp_no,name,DEPT_NM"
                      lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn_nm "
                      lgStrSQL = lgStrSQL & " From  HAA010T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 	
                      lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
                Case "N"
                      lgStrSQL = "Select TOP 1 mil_type,mil_kind,mil_start,mil_end,mil_grade,mil_branch,mil_no,emp_no,name,DEPT_NM"
                      lgStrSQL = lgStrSQL & ",dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn_nm "
                      lgStrSQL = lgStrSQL & " From  HAA010T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
                      lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC "
                      
             End Select
      Case "C"
      Case "U"
      Case "D"
    End Select
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
    End Select    
       
</Script>	
