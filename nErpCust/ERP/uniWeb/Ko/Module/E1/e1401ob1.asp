<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
        'Case "UID_M0002"       
        '    Call SubBizQuery2()
'             Call SubBizSaveSingleUpdate()
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
    Dim TotNotTax
    Dim issueno
    Dim issueno1
    dIM txtissueno1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")     ' 사번으로조회 
    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          lgPrevNext = ""
          Call SubBizQuery()
       End If
    Else

        Call CommonQueryRs(" co_full_nm, repre_nm "," b_company ", " co_cd=" & FilterVar("U2000", "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
%>
<Script Language=vbscript>
        With parent.frm1
            .txtEmp_no.Value = "<%=ConvSPChars(lgKeyStream(0))%>"
            .txtName.Value = "<%=ConvSPChars(lgObjRs("Name"))%>"
            .txtDept_nm.value = "<%=ConvSPChars(lgObjRs("DEPT_NM"))%>"    
            .txtroll_pstn.value = "<%=ConvSPChars(lgObjRs("roll_pstn"))%>"

            .txtEmp_no1.Value = "<%=ConvSPChars(lgKeyStream(0))%>"
            .txtName1.Value = "<%=ConvSPChars(lgObjRs("Name"))%>"

            .txtres_no.value = "<%=ConvSPChars(lgObjRs("res_no"))%>"
            .txtdomi.value = "<%=ConvSPChars(lgObjRs("domi"))%>"
            .txtaddr.value = "<%=ConvSPChars(lgObjRs("addr"))%>"
            .txtentr_dt.value = "<%=UNIDateClientFormat(lgObjRs("entr_dt"))%>"
            .txtretire_dt.value = "<%=UNIDateClientFormat(lgObjRs("retire_dt"))%>"

            .txtco_full_nm.value = replace("<%=ConvSPChars(lgF0)%>", Chr(11), "")
            .txtrepre_nm.value = replace("<%=ConvSPChars(lgF1)%>", Chr(11), "")
        End With
</Script>       
<%

        issueno = uniConvDateToYYYYMMDD(lgSvrDateTime,gServerDateFormat,"")
        Call CommonQueryRs(" issue_no1 "," HAA170T ", " issue_no= " & FilterVar(issueno, "''", "S") & " and proof_use=" & FilterVar("재직증명서", "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        if  lgF0 = "X" then
            issueno1 = "1"
            lgStrSQL = "INSERT INTO HAA170T (issue_no, issue_no1, proof_type, issue_dt,"
            lgStrSQL = lgStrSQL & " proof_use, emp_no, emp_name, print_emp_no )"
            lgStrSQL = lgStrSQL & " VALUES ("
            lgStrSQL = lgStrSQL & "  " & FilterVar(issueno, "''", "S") & ","
            lgStrSQL = lgStrSQL & "  " & FilterVar(issueno1, "''", "S") & ","
            lgStrSQL = lgStrSQL & " " & FilterVar("1", "''", "S") & ","           'PROOF TYPE
            lgStrSQL = lgStrSQL & "  " & FilterVar(lgSvrDateTime, "''", "S") & ","
            lgStrSQL = lgStrSQL & " " & FilterVar("재직증명서", "''", "S") & ","  'PROOF USE
            lgStrSQL = lgStrSQL & "  " & FilterVar(lgKeyStream(0), "''", "S") & ","
            lgStrSQL = lgStrSQL & "  " & FilterVar(ConvSPChars(lgObjRs("Name")), "''", "S") & ","
            lgStrSQL = lgStrSQL & "  " & FilterVar(lgKeyStream(0), "''", "S") & " )"
			
'			Response.Write lgStrSQL

            lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	        Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
        else
            issueno1 = Cint(lgF0) + 1
            lgStrSQL = "UPDATE HAA170T SET issue_no1= " & FilterVar(issueno1, "''", "S") & ""
            lgStrSQL = lgStrSQL & " WHERE issue_no= " & FilterVar(issueno, "''", "S") & ""
            lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	        Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
        end if

        txtissueno1 = issueno & "-" & issueno1
		%>
		<Script Language=vbscript>
		        With parent.frm1
		            .txtissueno.Value = "<%=txtissueno1%>"
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
    
    
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()


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

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pMode 
      Case "R"
        lgStrSQL = "Select emp_no, name, dept_nm, res_no, domi, addr,entr_dt,retire_dt," 
        lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) as roll_pstn "
        lgStrSQL = lgStrSQL & " From  haa010t "
        lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode
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
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
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
       '      Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	
