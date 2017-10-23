<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

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
    Dim iKey1,iRet, i, j
    Dim lgStrSQL1
    Dim Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt
	Dim tempArr(13,8)
	Dim cdArr(7)
    Dim str_yy,end_yy,str_mm,end_mm,for_month,for_yy
    Dim year_order,mon_order
'    On Error Resume Next                                                             '☜: Protect system from crashing
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
	lgStrSQL1 = "SELECT crt_strt_yy,crt_strt_mm,crt_end_yy,crt_end_mm FROM HDA150T WHERE allow_cd = " & FilterVar("Y01", "''", "S") 

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
		            parent.frm1.MONTH<%=i%>.value = "<%=uniConvDateAtoB(year_order & "-" & mon_order & "-" & "01", gServerDateFormat, gDateFormatYYYYMM)%>" 
			</Script>       
			<%		
			mon_order = mon_order +1
		next
 		
	end if
	
'column별 근태코드,근태명 셋팅 
	lgStrSQL1 = " select top 7 dilig_nm,dilig_cd ,a.cnt from hca010t ,( select count(*) cnt  from hca010t where dilig_cd in (select dilig_cd  From  HDA100T where flag = '2'  ) ) a where dilig_cd in (select dilig_cd  From  HDA100T where flag = '2') order by dilig_seq "
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL1,"X","X") = False Then
        Call SetErrorStatus()
    else 
  
		for i=1 to lgObjRs("cnt")
		   	cdArr(i) = lgObjRs("dilig_cd")
			%>
			<Script Language=vbscript>
		            parent.frm1.TITLE<%=i%>.value = "<%=ConvSPChars(lgObjRs("dilig_nm"))%>"
			</Script> 
			<%		
		   	lgObjRs.MoveNext
		next
 		
	end if

	iKey1 = FilterVar(ConvSPChars(emp_no), "''", "S")	
'	iKey1 = iKey1 & " AND ( dilig_dt between " & FilterVar(UNIGetFirstDay( str_yy & "-" & str_mm & "-01","YYYY-MM-DD"),"''","S") &" and " & FilterVar(UNIGetLastDay( end_yy & "-" & end_mm & "-01","YYYY-MM-DD"),"''","S")  & ")"
	iKey1 = iKey1 & " AND  convert(varchar(4),dilig_dt,112) = " & FilterVar(str_yy,"''","S")  

    Call SubMakeSQLStatements("R",iKey1)                                     '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Response.end
        Call SetErrorStatus()
    Else
'arrary에 count 셋팅 
        for i=1 to for_month+1
		    for j=1 to 7
				tempArr(i,j)= 0
			Next
        Next 

'arrary값 화면에 셋팅 
        Do While Not lgObjRs.EOF
			for i=1 to for_month		
				if UniConvDateToYYYYMM(tempArr(i,0),"YYYY-MM-DD","") = ConvSPChars(lgObjRs("yyyymm"))  then
					for j=1 to 7
						if ConvSPChars(lgObjRs("dilig_cd")) = cdArr(j) then
							tempArr(i,j)=lgObjRs("count")
							exit for				
						end if
					next
				end if
			next
			lgObjRs.MoveNext
        Loop 

		lgstrData = "" 

        for i=1 to for_month+1
		    for j=1 to 7
				lgstrData = lgstrData & Chr(11) & tempArr(i,j)
				if i= (for_month+1) then		    
					%>
					<Script Language=vbscript>
				            parent.frm1.SUM<%=j%>.value = "<%=tempArr((for_month+1),j)%>"
					</Script>       
					<%
				else
					tempArr(for_month+1,j) = 0 +tempArr(for_month+1,j) + tempArr(i,j)
				end if
			Next
			lgstrData = lgstrData & Chr(11) & Chr(12) 
        Next    
    End If
    ' Call svrmsgbox(lgstrData, vbinformation, i_mkscript)
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
			lgStrSQL = "SELECT  left(convert(varchar(8),DILIG_DT,112),6) yyyymm, DILIG_CD, count(*) count"
			lgStrSQL = lgStrSQL & " FROM HCA060T "
            lgStrSQL = lgStrSQL & " WHERE EMP_NO = " & pCode 	
            lgStrSQL = lgStrSQL & "   AND DILIG_CD in ( select top 7 dilig_cd from hca010t where dilig_cd in (select dilig_cd  From  HDA100T where flag = '2') order by dilig_seq ) "
            lgStrSQL = lgStrSQL & " GROUP BY left(convert(varchar(8),DILIG_DT,112),6), DILIG_CD"
            lgStrSQL = lgStrSQL & " ORDER BY yyyymm"
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
