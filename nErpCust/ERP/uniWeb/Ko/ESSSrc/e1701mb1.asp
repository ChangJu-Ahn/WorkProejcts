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
                                                               '☜: Hide Processing message
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection


    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '☜: Save,Update
             Call SubBizSave()
        Case "UID_M0003"
             Call SubBizDelete()
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
    iRet = EmpBaseDiligAuthCheck(lgKeyStream(0),lgKeyStream(4),lgKeyStream(5),lgKeyStream(6),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
 
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
       With Parent
			.FncNew()       
       End With         
</Script> 
<%		
			Call DisplayMsgBox("800454", vbInformation, "", "", I_MKSCRIPT)	
		end if
		Response.End
	end if	
	
		
    if  lgPrevNext = "N" or lgPrevNext = "P" then
		%>
		<Script Language=vbscript>
		       With Parent
					.FncNew()
		       End With         
		</Script>       
		<%   
	        Call SetErrorStatus()
	else 		
	    iKey1 = FilterVar(ConvSPChars(emp_no), "''", "S")
	    iKey1 = iKey1 & " AND trip_strt_dt  =  " & FilterVar( lgKeyStream(1), "''", "S") & ""
		iKey1 = iKey1 & " AND trip_cd  =  " & FilterVar(lgKeyStream(3), "''", "S") & ""

	    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements
	    
	    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	        lgStrPrevKeyIndex = ""
	%>
	<Script Language=vbscript>
	       With Parent
	       
	            .frm1.txttrip_cd.value = ""
	            .frm1.txttrip_loc.value = ""
	            .frm1.txtremark.value = ""
	            .frm1.txttrip_amt.value = ""
	            .frm1.txtapp_emp_no.value = ""
	            .frm1.txtapp_name.value = ""
	       End With         
	</Script>       
	<%            
	        Call SetErrorStatus()
	    Else

	%>
	<Script Language=vbscript>
		       With Parent	
       
		            .frm1.txttrip_strt_dt.value = "<%=UNIDateClientFormat(lgObjRs("trip_strt_dt"))%>"
		            .frm1.txttrip_end_dt.value = "<%=UNIDateClientFormat(lgObjRs("trip_end_dt"))%>"
		            .frm1.txttrip_cd.value = "<%=ConvSPChars(lgObjRs("trip_cd"))%>"
		            .frm1.txttrip_loc.value = "<%=ConvSPChars(lgObjRs("trip_loc"))%>"
		            .frm1.txtremark.value = "<%=ConvSPChars(lgObjRs("remark"))%>"
		            .frm1.txttrip_amt.value = "<%=UNINumClientFormat(lgObjRs("trip_amt"), ggAmtOfMoney.DecPoint, 0)%>"
		            .frm1.txtapp_emp_no.value = "<%=ConvSPChars(lgObjRs("app_emp_no"))%>"
		            .frm1.txtapp_name.value = "<%=ConvSPChars(lgObjRs("app_name"))%>"
		       End With        
	</Script>       
	<%     
	    End If
	End If
    Call SubCloseRs(lgObjRs)

End Sub


'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

	Dim trip_start_dt, trip_end_dt

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	trip_start_dt = UNIConvDateCompanyToDB((lgKeyStream(1)),NULL)
	trip_end_dt   = UNIConvDateCompanyToDB((lgKeyStream(2)),NULL)

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
              
			'출장근태(e11080t)를 check 한다 
			Call CommonQueryRs(" isnull(count(emp_no),0) "," e11080t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND ((TRIP_STRT_DT between  " & FilterVar(trip_start_dt, "''", "S") & " AND  " & FilterVar(trip_end_dt, "''", "S") & ") OR (TRIP_END_DT  between  " & FilterVar(trip_start_dt, "''", "S") & " AND  " & FilterVar( trip_end_dt, "''", "S") & "))" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
						
			If Trim(Replace(lgF0,Chr(11),"")) = 0 then
			Else
                Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
                lgErrorStatus = "YES"
			    Exit sub                                    '바로 return한다 
			End if
			
			'기간근태(e11070t)에서 기간에 (중복일자)속했는지를 check 한다 
			Call CommonQueryRs(" isnull(count(emp_no),0) "," e11070t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND ((dilig_strt_dt between  " & FilterVar( trip_start_dt, "''", "S") & " AND  " & FilterVar(trip_end_dt, "''", "S") & ") OR (dilig_end_dt  between  " & FilterVar(trip_start_dt, "''", "S") & " AND  " & FilterVar(trip_end_dt, "''", "S") & "))" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
						
			If Trim(Replace(lgF0,Chr(11),"")) = 0 then
			Else
                Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
                lgErrorStatus = "YES"
			    Exit sub                                    '바로 return한다 
			End if
			
			'기간근태(hca050t)에서 기간에 (중복일자)속했는지를 check 한다. 만약 없으면 일일근태(hca060t)에 있는지도 check 한다.
			Call CommonQueryRs(" isnull(count(emp_no),0) "," hca050t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND ((dilig_strt_dt between  " & FilterVar( trip_start_dt, "''", "S") & " AND  " & FilterVar(trip_end_dt, "''", "S") & ") OR (dilig_end_dt  between  " & FilterVar(trip_start_dt, "''", "S") & " AND  " & FilterVar(trip_end_dt, "''", "S") & "))" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
						
			If Trim(Replace(lgF0,Chr(11),"")) = 0 then
			    Call CommonQueryRs(" isnull(count(emp_no),0) "," hca060t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND (dilig_dt between  " & FilterVar(trip_start_dt, "''", "S") & " AND  " & FilterVar(trip_end_dt, "''", "S") & ")" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
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
				    strClose_dt = Trim(Replace(lgF1,Chr(11),""))

				    IF strClose_type = "1" THEN 
				    	strClose_dt = UNIDateAdd("D",-1,strClose_dt,gServerDateFormat)
				    END IF 

				    IF (uniConvDateToYYYYMMDD(lgKeyStream(1),gDateFormat,"-") > uniConvDateToYYYYMMDD(strClose_dt,gServerDateFormat,"-")) AND (uniConvDateToYYYYMMDD(strClose_dt,gServerDateFormat,"-") < uniConvDateToYYYYMMDD(lgKeyStream(2),gDateFormat,"-")) THEN 
				    ELSE
				    dim tttt
                        Call DisplayMsgBox("800291", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
                        lgErrorStatus = "YES"
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
				    strClose_dt = uniConvDateToYYYYMMDD(Trim(Replace(lgF1,Chr(11),"")),gServerDateFormat,"-")
			                
				    IF strClose_type = "1" THEN 
				    	strClose_dt = UNIDateAdd("D",-1,strClose_dt,gServerDateFormat)
				    END IF 
				    IF (uniConvDateToYYYYMMDD(lgKeyStream(1),gDateFormat,"-") > strClose_dt) AND (strClose_dt < uniConvDateToYYYYMMDD(lgKeyStream(2),gDateFormat,"-")) THEN 
				    ELSE
                        Call DisplayMsgBox("800291", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
                        lgErrorStatus = "YES"
				        Exit sub                                    '바로 return한다 
				    END IF 	 
				Next
			End if	                
            Call SubBizSaveSingleUpdate()
    
    End Select

End Sub	

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  E11080T"
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(lgKeyStream(0), "''", "S")                              ' 사번char(10)
       	lgStrSQL = lgStrSQL & "   AND trip_strt_dt = " & FilterVar(UNIConvDateCompanyToDB(lgKeyStream(2),NULL),"NULL","S")    

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO E11080T("
    lgStrSQL = lgStrSQL & " emp_no, "
    lgStrSQL = lgStrSQL & " trip_strt_dt, "
    lgStrSQL = lgStrSQL & " trip_end_dt, "
    lgStrSQL = lgStrSQL & " trip_cd, "
    lgStrSQL = lgStrSQL & " trip_loc, "
    lgStrSQL = lgStrSQL & " remark, "
    lgStrSQL = lgStrSQL & " trip_amt, "
    lgStrSQL = lgStrSQL & " app_emp_no, "
    lgStrSQL = lgStrSQL & " isrt_emp_no, "
    lgStrSQL = lgStrSQL & " updt_emp_no ) "
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEmp_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txttrip_strt_dt"),NULL),"NULL","S")   & "," 
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txttrip_end_dt"),NULL),"NULL","S")   & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txttrip_cd"), "''", "S") & ", '"
    lgStrSQL = lgStrSQL & Request("txttrip_loc") & "','"
    lgStrSQL = lgStrSQL & Request("txtremark") & "',"
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txttrip_amt"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtapp_emp_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtapp_emp_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtapp_emp_no"), "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
End Sub

		
'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  E11080T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " trip_end_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txttrip_end_dt"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " trip_cd = " & FilterVar(lgKeyStream(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & " trip_loc =  " & FilterVar(Request("txttrip_loc"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " remark =  " & FilterVar(Request("txtremark"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " trip_amt = " & UNIConvNum(Request("txttrip_amt"),0) & ","
    lgStrSQL = lgStrSQL & " app_emp_no = " & FilterVar(Request("txtapp_emp_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " updt_emp_no = " & FilterVar(Request("txtEmp_no"), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no =  " & FilterVar(Request("txtEmp_no"), "''", "S") & ""
    lgStrSQL = lgStrSQL & "   AND trip_strt_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txttrip_strt_dt"),NULL),"NULL","S") 

'Response.Write lgStrSQL
'Response.End
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount

    Select Case pMode 
      Case "R"
              lgStrSQL = "Select emp_no,trip_strt_dt,trip_end_dt,trip_cd,trip_loc,remark,trip_amt,app_emp_no, dbo.ufn_H_GetEmpName(app_emp_no)  as app_name " 
              lgStrSQL = lgStrSQL & " From  E11080T "
        	  lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode
'Response.Write lgStrSQL
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
             Parent.DBSaveOk("<%=lgKeyStream(3)%>")
          Else
             Parent.DBSaveFail
             'Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If
       Case "UID_M0003"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If
    End Select
       
</Script>	
