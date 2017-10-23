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

    Call HideStatusWnd_uniSIMS
    Dim emp_no 
    Dim name
    Dim dept_nm
                                                               
    lgErrorStatus     = "NO"
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
    Dim iKey1,iRet
    Dim lgStrSQL1
    Dim Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt
	Dim tempArr(200,8),cdArr(7)
    Dim str_yy,end_yy,str_mm,end_mm,for_month,for_yy
    Dim year_order,mon_order
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iRet = EmpBaseDiligAuthCheck(lgKeyStream(0),lgKeyStream(3),lgKeyStream(4),lgKeyStream(1),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)

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
'연차수당기준등록에 있는 연차발생분 날짜를 가져옴 
	lgStrSQL1 = "SELECT crt_strt_yy,crt_strt_mm,crt_end_yy,crt_end_mm FROM HDA150T WHERE allow_cd = " & FilterVar("Y01", "''", "S") & ""
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL1,"X","X") = False Then
        Call SetErrorStatus()
    else 
		str_yy = 0 + lgKeyStream(2)
		for_yy = (0+ConvSPChars(lgObjRs("crt_end_yy")) - ConvSPChars(lgObjRs("crt_strt_yy")))
		end_yy = 0 + str_yy +  for_yy

		str_mm = cint(ConvSPChars(lgObjRs("crt_strt_mm")))
		end_mm = cint(ConvSPChars(lgObjRs("crt_end_mm")))
		for_month = 12*for_yy + end_mm - str_mm + 1

		year_order = str_yy
		mon_order = str_mm
		for i=1 to for_month
			if ((str_mm+i-1) mod 13)=0 then
				year_order = year_order +1
				mon_order = mon_order -12
	
			end if
			tempArr(i,0) =  year_order & "-" & mon_order & "-" & "01"
%>
<Script Language=vbscript>
        With parent.frm1
            .MONTH<%=i%>.value = "<%=uniConvDateAtoB(year_order & "-" & mon_order & "-" & "01", gServerDateFormat, gDateFormatYYYYMM)%>" 
        End With 
</Script>       
<%		
			mon_order = mon_order +1
		next

	end if

'column별 근태코드,근태명 셋팅 
	lgStrSQL1 = " select top 7 dilig_nm,dilig_cd from hca010t order by dilig_seq "
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL1,"X","X") = False Then
        Call SetErrorStatus()
    else 
		for i=1 to 7
		   	cdArr(i) = lgObjRs("dilig_cd")
%>
<Script Language=vbscript>
        With parent.frm1
            .TITLE<%=i%>.value = "<%=ConvSPChars(lgObjRs("dilig_nm"))%>"
        End With 
</Script>       
<%		
		   	lgObjRs.MoveNext
		next
	end if

	iKey1 = FilterVar(ConvSPChars(emp_no), "''", "S")				
	iKey1 = iKey1 & " AND ( dilig_dt > " & FilterVar(str_yy & "-" & str_mm & "-01' and dilig_dt<'" & UNIGetLastDay( end_yy & "-" & end_mm & "-01","YYYY-MM-DD"), "''", "S") & ")"
	
    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Response.end
        Call SetErrorStatus()
    Else
        for i=1 to for_month+1
		    for j=1 to 7
				tempArr(i,j)= 0
			Next
        Next    
'arrary에 count 셋팅 

        Do While Not lgObjRs.EOF

			for i=1 to for_month
				if UniConvDateToYYYYMM(tempArr(i,0),"YYYY-MM-DD","") = UniConvDateToYYYYMM(lgObjRs("year") & "-"&lgObjRs("month") & "-01","YYYY-MM-DD","")  then
					for j=1 to 7
						if lgObjRs("dilig_cd") = cdArr(j) then
							tempArr(i,j)=lgObjRs("count")
							exit for				
						end if
					next
				end if
			next
			lgObjRs.MoveNext			
        Loop 
'arrary값 화면에 셋팅 
				
		lgstrData = ""    	
        for i=1 to for_month+1
		    for j=1 to 7

				lgstrData = lgstrData & Chr(11) & tempArr(i,j)
				if i= (for_month+1) then		    
%>
<Script Language=vbscript>
        With parent.frm1
            .SUM<%=j%>.value = "<%=tempArr((for_month+1),j)%>"
        End With 
</Script>       
<%
				else
					tempArr(for_month+1,j) = 0 +tempArr(for_month+1,j) + tempArr(i,j)
				end if
			Next
			lgstrData = lgstrData & Chr(11) & Chr(12) 
        Next    
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

			lgStrSQL = "SELECT   year(DILIG_DT) year,month(DILIG_DT) month, DILIG_CD,count(isrt_dt) count"
			lgStrSQL = lgStrSQL & " FROM HCA060T "
            lgStrSQL = lgStrSQL & " WHERE EMP_NO = " & pCode 	
            lgStrSQL = lgStrSQL & " and dilig_cd in ( select top 7 dilig_cd from hca010t order by dilig_seq ) "
            lgStrSQL = lgStrSQL & " GROUP BY emp_no,year(DILIG_DT),month(DILIG_DT),DILIG_CD"
            lgStrSQL = lgStrSQL & " ORDER BY month"
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
       Case "UID_M0002"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
    End Select    
       
</Script>	
