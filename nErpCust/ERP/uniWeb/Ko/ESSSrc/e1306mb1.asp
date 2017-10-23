<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd_uniSIMS
    Dim emp_no
    Dim name
    Dim dept_nm
                                                               '☜: Hide Processing message
    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        
    
    Call SubOpenDB(lgObjConn)															'☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"																'☜: Query
             Call SubBizQuery()
    End Select
    
    Call SubCloseDB(lgObjConn)															'☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1
    Dim Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    call  SubEmpBase(lgKeyStream(0),"1",lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
 
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
	if emp_no = "" then
        lgErrorStatus = "YES"
        exit sub
    end if 
	 iKey1 = FilterVar(lgKeyStream(0),"''", "S")
	 iKey1 = iKey1 & " and YEAR_YY  = "  &  FilterVar(lgKeyStream(2),"'%'", "S")	 
	 iKey1 = iKey1 & " AND  MED_DT >=  " & FilterVar(lgKeyStream(2) & "-01-01", "''", "S")  
'	 iKey1 = iKey1 & " AND internal_cd LIKE '" &FilterVar(lgKeyStream(3),"'%'", "S")
'	 iKey1 = iKey1 & " AND retire_dt is null"    

    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
    
        lgstrData = ""
        iDx       = 1  
        Do While Not lgObjRs.EOF
            emp_no = ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MED_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MED_NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MED_RGST_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_REL"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_REL_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_RES_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MED_AMT"), ggAmtOfMoney.DecPoint, 0) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MED_TEXT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUBMIT_FLAG"))
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
                    lgStrSQL = "Select MED_DT, MED_AMT, MED_TEXT, MED_NAME,MED_RGST_NO,FAMILY_NM,FAMILY_REL, dbo.ufn_GetCodeName('H0140',FAMILY_REL) FAMILY_REL_NM"
                    lgStrSQL = lgStrSQL & " , FAMILY_RES_NO,FAMILY_TYPE,CASE FAMILY_TYPE WHEN 'A' THEN '장애자' WHEN 'B' THEN '경로자' ELSE '' END FAMILY_TYPE_NM ,submit_flag "
                    lgStrSQL = lgStrSQL & " From HFA130T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode

              End Select
    End Select
'    Response.Write lgStrSQL
'    Response.End
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
