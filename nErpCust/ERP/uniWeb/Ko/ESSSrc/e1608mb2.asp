<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

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
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    Call SubOpenDB(lgObjConn)															'¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"																'¢Ð: Query
             Call SubBizQuery()
    End Select
    
    Call SubCloseDB(lgObjConn)															'¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1,iRet, i
    Dim Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt
    Dim dilig_cd,dbdate
    Dim lgStrSQL1,cdArr(7)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
	else 
   	    
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

	lgStrSQL1 = " select top 7 dilig_nm,dilig_cd ,a.cnt from hca010t ,( select count(*) cnt  from hca010t where dilig_cd in (select dilig_cd  From  HDA100T where flag = '2'  ) ) a where dilig_cd in (select dilig_cd  From  HDA100T where flag = '2') order by dilig_seq "
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL1,"X","X") = False Then
        Call SetErrorStatus()
    else 
		for i=1 to lgObjRs("cnt")
		   	cdArr(i) = lgObjRs("dilig_cd")
		   	lgObjRs.MoveNext
		next    
	end if
	for i=1 to lgObjRs("cnt")
	   	cdArr(i) = lgObjRs("dilig_cd")
	   	lgObjRs.MoveNext
	next
	dilig_cd = cdArr(lgKeyStream(3))

	dbdate = uniConvDateAtoB(lgKeyStream(2), gDateFormatYYYYMM, gServerDateFormat)
	iKey1 = FilterVar(ConvSPChars(emp_no), "''", "S")
	iKey1 = iKey1 & " AND year(dilig_dt)=year(" & FilterVar(dbdate, "''", "S") & ")"
	iKey1 = iKey1 & " AND month(dilig_dt)=month(" & FilterVar(dbdate, "''", "S") & ")"
	iKey1 = iKey1 & " AND dilig_cd = " & FilterVar(dilig_cd, "''", "S")	

    Call SubMakeSQLStatements("R",iKey1)                                     '¢Ð : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
    
        lgstrData = ""
        iDx       = 1

        Do While Not lgObjRs.EOF
'			if iDx=1 then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_nm"))
'			else
'				lgstrData = lgstrData & Chr(11)
'			end if
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("dilig_dt"))
            if ConvSPChars(lgObjRs("name")) = "" then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("isrt_emp_no"))
			else
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
			end if
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("isrt_emp_no"))
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
					lgStrSQL = "SELECT dilig_dt, dbo.ufn_H_GetCodeName(" & FilterVar("hca010t", "''", "S") & ",dilig_cd,'') dilig_nm"
					lgStrSQL = lgStrSQL & " ,isrt_emp_no,dbo.ufn_H_GetEmpName(isrt_emp_no) name"
					lgStrSQL = lgStrSQL & " FROM HCA060T  "
                    lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode 	
                    lgStrSQL = lgStrSQL & " ORDER BY dilig_dt DESC "
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.End                    
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
       Case "UID_M0002"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
    End Select    
       
</Script>	
