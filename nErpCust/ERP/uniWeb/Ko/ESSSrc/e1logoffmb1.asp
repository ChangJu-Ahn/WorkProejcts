<%@ LANGUAGE=VBSCript%>

<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/uniSimsClassID.inc" -->

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>


<%
Dim strDBCon

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd_uniSIMS

    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '☜: Save,Update
            Call SubBizSave()
        Case "UID_M0003"
             Call SubBizDelete()
    End Select
  

'============================================================================================================
' Name : SubBizQuery
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1
    Dim DiligAuth
    Dim Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   iRet = SubEmpBaseDiligAuth(lgKeyStream(0),lgKeyStream(1),lgKeyStream(5),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
    If iRet = True Then
%>
        <Script Language=vbscript>
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
            end if
            Response.End
    End If

    if  lgKeyStream(1) = "" then 
        lgErrorStatus = "YES"
        exit sub
    end if

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")

    if lgKeyStream(2) = "" or lgKeyStream(3) = "" then 
        lgErrorStatus = "NO"
        return
    end if
    iKey1 = iKey1 & " AND ((dilig_strt_dt between  " & FilterVar( lgKeyStream(2), "''", "S") & " AND  " & FilterVar( lgKeyStream(3), "''", "S") & ") OR (dilig_end_dt  between  " & FilterVar( lgKeyStream(2), "''", "S") & " AND  " & FilterVar( lgKeyStream(3), "''", "S") & "))"
    iKey1 = iKey1 & " AND dilig_cd =  " & FilterVar(lgKeyStream(4), "''", "S") & ""
    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call SetErrorStatus()
    Else
%>
<Script Language=vbscript>
       With Parent
            .frm1.txtdilig_strt_dt.value = "<%=ConvSPChars(lgObjRs("dilig_strt_dt"))%>"
            .frm1.txtdilig_end_dt.value = "<%=ConvSPChars(lgObjRs("dilig_end_dt"))%>"
            .frm1.txtdilig_cd.value = "<%=ConvSPChars(lgObjRs("dilig_cd"))%>"
            .frm1.txtremark.value = "<%=ConvSPChars(lgObjRs("remark"))%>"
            .frm1.txtapp_emp_no.value = "<%=ConvSPChars(lgObjRs("app_emp_no"))%>"
            .frm1.txtapp_name.value = "<%=ConvSPChars(lgObjRs("app_name"))%>"
       End With          
</Script>       
<%     
    End If
    Call SubCloseRs(lgObjRs)
End Sub    


'============================================================================================================
' Name : SubBizSave
'============================================================================================================
Sub SubBizSave()
	Dim counts
	Dim i
	Dim strInput_emp_no
	Dim strClose_type
	Dim strClose_dt
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    
	Call CommonQueryRs(" name "," haa010t ","  emp_no =  " & FilterVar(Request("txtapp_emp_no"), "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  lgF0 = "X" then
        Call DisplayMsgBox("800094", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        lgErrorStatus = "YES"
%>
<Script Language=vbscript>
        parent.frm1.txtApp_emp_no.focus()
</Script>       
<%     
        exit sub
    end if

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
 
   
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
			'기간근태(e11070t)에서 기간에 (중복일자)속했는지를 check 한다 
			Call CommonQueryRs(" isnull(count(emp_no),0) "," e11070t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND ((dilig_strt_dt between  " & FilterVar( lgKeyStream(2), "''", "S") & " AND  " & FilterVar( lgKeyStream(3), "''", "S") & ") OR (dilig_end_dt  between  " & FilterVar( lgKeyStream(2), "''", "S") & " AND  " & FilterVar( lgKeyStream(3), "''", "S") & "))" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If Trim(Replace(lgF0,Chr(11),"")) = 0 then
			Else
		        Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
                lgErrorStatus = "YES"
			    Exit sub                                    '바로 return한다 
			End if
			
			'기간근태(hca050t)에서 기간에 (중복일자)속했는지를 check 한다. 만약 없으면 일일근태(hca060t)에 있는지도 check 한다.
			Call CommonQueryRs(" isnull(count(emp_no),0) "," hca050t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND ((dilig_strt_dt between  " & FilterVar( lgKeyStream(2), "''", "S") & " AND  " & FilterVar( lgKeyStream(3), "''", "S") & ") OR (dilig_end_dt  between  " & FilterVar( lgKeyStream(2), "''", "S") & " AND  " & FilterVar( lgKeyStream(3), "''", "S") & "))" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			If Trim(Replace(lgF0,Chr(11),"")) = 0 then
			    Call CommonQueryRs(" isnull(count(emp_no),0) "," hca060t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND (dilig_dt between  " & FilterVar( lgKeyStream(2), "''", "S") & " AND  " & FilterVar( lgKeyStream(3), "''", "S") & ")" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			    If Trim(Replace(lgF0,Chr(11),"")) > 0 then
			        Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
                    lgErrorStatus = "YES"
			        Exit sub                                    '바로 return한다.
			    End if
			Else
		        Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
                lgErrorStatus = "YES"
			    Exit sub                                    '바로 return한다 
			End if

		    
		    '근태마감을 체크한다.
		    Call CommonQueryRs(" close_type, close_dt, emp_no, COUNT(close_dt) as counts "," hda270t ","  ORG_CD = " & FilterVar("1", "''", "S") & " AND PAY_GUBUN = " & FilterVar("Z", "''", "S") & " AND PAY_TYPE  = " & FilterVar("#", "''", "S") & "   GROUP BY emp_no,close_type,close_dt" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
			If Trim(Replace(lgF3,Chr(11),""))="" or Trim(Replace(lgF3,Chr(11),""))="X" Then
			Else
				counts = Trim(Replace(lgF3,Chr(11),""))
				For i = 1 to counts
				    strInput_emp_no = Trim(Replace(lgF2,Chr(11),""))
				    strClose_type = Trim(Replace(lgF0,Chr(11),""))
				    strClose_dt = CDate(Trim(Replace(lgF1,Chr(11),"")))
			                
				    IF strClose_type = "1" THEN 
				    	strClose_dt = strClose_dt - 1
				    END IF 

				    IF (CDate(lgKeyStream(2)) > CDate(strClose_dt)) AND (CDate(strClose_dt) < CDate(lgKeyStream(3))) THEN 
				    ELSE
				        Call DisplayMsgBox("800291", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				        Exit sub                                    '바로 return한다 
				    END IF 	 
				Next
			End if	                
			Call SubBizSaveSingleCreate()  
        
        Case  OPMD_UMODE                                                             '☜ : Update
		    '근태마감을 체크한다.
		    Call CommonQueryRs(" close_type, close_dt, emp_no, COUNT(close_dt) as counts "," hda270t ","  ORG_CD = " & FilterVar("1", "''", "S") & " AND PAY_GUBUN = " & FilterVar("Z", "''", "S") & " AND PAY_TYPE  = " & FilterVar("#", "''", "S") & "   GROUP BY emp_no,close_type,close_dt" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
			If Trim(Replace(lgF3,Chr(11),""))="" or Trim(Replace(lgF3,Chr(11),""))="X" Then
			Else
				counts = Trim(Replace(lgF3,Chr(11),""))
				For i = 1 to counts
				    strInput_emp_no = Trim(Replace(lgF2,Chr(11),""))
				    strClose_type = Trim(Replace(lgF0,Chr(11),""))
				    strClose_dt = CDate(Trim(Replace(lgF1,Chr(11),"")))
			                
				    IF strClose_type = "1" THEN 
				    	strClose_dt = strClose_dt - 1
				    END IF 

				    IF (CDate(lgKeyStream(2)) > CDate(strClose_dt)) AND (CDate(strClose_dt) < CDate(lgKeyStream(3))) THEN 
				    ELSE
				        Call DisplayMsgBox("800291", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				        Exit sub                                    '바로 return한다 
				    END IF 	 
				Next
			End if	                
            Call SubBizSaveSingleUpdate()
    
    End Select

End Sub	

'============================================================================================================
' Name : SubBizDelete
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Dim pObjConn,strCompany,Connect,pSource,pRs,strSAPwd
'	Dim Pid
      
			strSAPwd =  Request.Cookies("unierp")("gSAPwd")
            strDBCon = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=" _
                        & strSAPwd & ";Initial Catalog=" & gDataBase & ";Data Source=" & gDBServer

       	Set pObjConn = Server.CreateObject("ADODB.Connection")
        	'pObjConn.ConnectionString  = strDBCon
        	pObjConn.ConnectionString  = gADODBConnString ' 수정하지 마세요 by lsm
      	    pObjConn.ConnectionTimeout = 60
       	    pObjConn.Open 

           pSource = "select spid from master.dbo.sysprocesses where program_name = 'uniSIMS-" & gCompanyNm & "' and hostname='" & gUsrId & "'"
           Set pRs = Server.CreateObject("ADODB.Recordset") 
            pRs.Open pSource,pObjConn,adOpenForwardOnly,adLockReadOnly 
           If pRs(0) <> 0 Then
             pSource = "Kill " & pRs(0) 
		     pObjConn.Execute pSource,,adCmdText+adExecuteNoRecords
              
           End If
           pRs.Close
           pObjConn.Close
           Set pRs = nothing
           Set pObjConn = nothing
           
           Response.Cookies("unierp")("gEmpNo")  = ""
           Response.Cookies("unierp")("gjdoiwp") = ""
           Response.Redirect "../uniSims.asp"
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO E11070T("
    lgStrSQL = lgStrSQL & " emp_no, "
    lgStrSQL = lgStrSQL & " dilig_strt_dt, "
    lgStrSQL = lgStrSQL & " dilig_end_dt, "
    lgStrSQL = lgStrSQL & " dilig_cd, "
    lgStrSQL = lgStrSQL & " remark, "
    lgStrSQL = lgStrSQL & " app_emp_no, "
    lgStrSQL = lgStrSQL & " isrt_emp_no, "
    lgStrSQL = lgStrSQL & " updt_emp_no ) "
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEmp_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtdilig_strt_dt"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtdilig_end_dt"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtdilig_cd"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtremark"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtapp_emp_no")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEmp_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEmp_no"), "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
	
End Sub


'============================================================================================================
' Name : SubBizSaveSingleUpdate
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgStrSQL = "UPDATE  E11070T"
    lgStrSQL = lgStrSQL & " SET   dilig_end_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtdilig_end_dt"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & "       remark = "        & FilterVar(UCase(Request("txtremark")), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       app_emp_no = "    & FilterVar(UCase(Request("txtapp_emp_no")), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       updt_emp_no = "   & FilterVar(Request("txtEmp_no"), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no = "        & FilterVar(Request("txtEmp_no"), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_strt_dt = " & FilterVar(Request("txtdilig_strt_dt"), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = "      & FilterVar(Request("txtdilig_cd"), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount
    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
                Case ""
                      lgStrSQL = "Select top 1 emp_no,dilig_strt_dt,dilig_cd,dilig_end_dt,app_emp_no,remark," 
                      lgStrSQL = lgStrSQL & " (select haa010t.name from haa010t where haa010t.emp_no = E11070T.app_emp_no) as app_name"
                      lgStrSQL = lgStrSQL & " From  E11070T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode 	
                Case "P"
                      lgStrSQL = "Select TOP 1 emp_no,dilig_strt_dt,dilig_cd,dilig_end_dt,app_emp_no,remark,"
                      lgStrSQL = lgStrSQL & " (select haa010t.name from haa010t where haa010t.emp_no = E11070T.app_emp_no) as app_name"
                      lgStrSQL = lgStrSQL & " From  E11070T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 	
                      lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
                Case "N"
                      lgStrSQL = "Select TOP 1 emp_no,dilig_strt_dt,dilig_cd,dilig_end_dt,app_emp_no,remark,"
                      lgStrSQL = lgStrSQL & " (select haa010t.name from haa010t where haa010t.emp_no = E11070T.app_emp_no) as app_name"
                      lgStrSQL = lgStrSQL & " From  E11070T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
                      lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC "
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
