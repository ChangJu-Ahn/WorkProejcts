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
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    Call SubOpenDB(lgObjConn)                                                        '☜: trip a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1
    Dim DiligAuth
    Dim Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt
    dim DiligAuths,strDiligAuth,login_emp_no,top_emp_no
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    iRet = EmpBaseDiligAuthCheck(lgKeyStream(0),lgKeyStream(4),lgKeyStream(5),lgKeyStream(1),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
 
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
		if  lgPrevNext = "N" then
			Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)
		elseif lgPrevNext = "P" then
			Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)
		else 
%>

<Script Language=vbscript>
        With parent.parent
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
			Call DisplayMsgBox("800454", vbInformation, "", "", I_MKSCRIPT)	
		end if
		Response.End
	end if	

	iKey1 = FilterVar(ConvSPChars(emp_no), "''", "S")
	iKey1 = iKey1 & "   AND ((trip_strt_dt between  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(2)),NULL), "''", "S") & " and  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(3)),NULL), "''", "S") & ")"
	iKey1 = iKey1 & "    OR  (trip_end_dt between  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(2)),NULL), "''", "S") & " and  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(3)),NULL), "''", "S") & "))"

    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements
    'response.write lgStrSQL
    'response.end
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
    
        lgstrData = ""
        iDx       = 1

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("trip_strt_dt"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("trip_end_dt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("trip_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("trip_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("trip_loc"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("trip_amt"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("remark"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("app_emp_no")))
            Select Case ConvSPChars(lgObjRs("app_yn"))
				Case "Y"
					lgstrData = lgstrData & Chr(11) & "승인"
				Case "R"
					lgstrData = lgstrData & Chr(11) & "반려"
				Case "N"
					lgstrData = lgstrData & Chr(11) & "미처리"
            End Select

            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 
        
    End If
    Call SubCloseRs(lgObjRs)

End Sub    

'============================================================================================================
' Name : SubtripSQLStatements
' Desc : trip SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount

    Select Case pMode 
      Case "R"
'             Select Case  lgPrevNext 
'                Case ""
                    lgStrSQL = "Select emp_no,trip_strt_dt, trip_cd,trip_end_dt,"
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetCodeName(" & FilterVar("hca010t", "''", "S") & ",trip_cd,'') as trip_nm , "                    
                    lgStrSQL = lgStrSQL & " trip_loc,remark,trip_amt, "
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(app_emp_no)  as app_emp_no,app_yn "
					lgStrSQL = lgStrSQL & " From E11080T "                    
                    lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode 
                    lgStrSQL = lgStrSQL & " ORDER BY trip_strt_dt DESC "

'                Case "P"
'                    lgStrSQL = "Select emp_no,trip_strt_dt, trip_cd,trip_end_dt,"
'                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetCodeName('hca010t',trip_cd,'') as trip_nm , "                    
'                    lgStrSQL = lgStrSQL & " trip_loc,remark,trip_amt, "
'                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(app_emp_no)  as app_emp_no,app_yn "
'					lgStrSQL = lgStrSQL & " From E11080T "                    
'                    lgStrSQL = lgStrSQL & " WHERE emp_no < " 	& pCode
'                    lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
'                Case "N"
'                    lgStrSQL = "Select emp_no,trip_strt_dt, trip_cd,trip_end_dt,"
'                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetCodeName('hca010t',trip_cd,'') as trip_nm , "                    
'                    lgStrSQL = lgStrSQL & " trip_loc,remark,trip_amt, "
'                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(app_emp_no)  as app_emp_no,app_yn "
'					lgStrSQL = lgStrSQL & " From E11080T "                    
'                    lgStrSQL = lgStrSQL & " WHERE emp_no > " 	& pCode
'                    lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC "
'             End Select
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
