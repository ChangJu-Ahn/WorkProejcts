<%@ LANGUAGE=VBSCript%>
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
    Dim emp_no
    Dim name
    Dim dept_nm
                                                            '¢Ð: Hide Processing message
    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        
	lgSvrDateTime = GetSvrDateTime
	    
    Call SubOpenDB(lgObjConn)															'¢Ð: Make a DB Connection
  
    Select Case lgOpModeCRUD
        Case "UID_M0001"																'¢Ð: Query
             Call SubBizQuery()
        Case "UID_M0002"																'¢Ð: Save,Update
             Call SubBizSaveSingleUpdate()
             Call SubBizSaveMulti()
        Case "UID_M0003"															'¢Ð: Delete
             Call SubCreatData()
    End Select
    
    Call SubCloseDB(lgObjConn)															'¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1
    Dim Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
	 iKey1 = iKey1 &" and YEAR_YY = "  &  FilterVar(lgKeyStream(2),"'%'", "S")	 
 
    Call SubMakeSQLStatements("R",iKey1)                                       '¢Ð : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
    
        lgstrData = ""
        iDx       = 1
        Do While Not lgObjRs.EOF
            emp_no = ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_REL"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_REL_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_RES_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASE_YN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PARIA_YN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CHILD_YN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INSUR_YN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MEDI_YN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EDU_YN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CARD_YN"))

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
                    lgStrSQL = "Select FAMILY_NAME, FAMILY_REL , dbo.ufn_GetCodeName('H0140',FAMILY_REL) as FAMILY_REL_NM, "
                    lgStrSQL = lgStrSQL & "  FAMILY_RES_NO , BASE_YN, PARIA_YN, CHILD_YN, INSUR_YN, MEDI_YN, EDU_YN, CARD_YN"
                    lgStrSQL = lgStrSQL & " From HFA150T"
                    lgStrSQL = lgStrSQL & " WHERE EMP_NO = " & pCode
                    lgStrSQL = lgStrSQL & " ORDER BY FAMILY_REL, FAMILY_RES_NO ASC"    
 
             End Select
    End Select
'Response.Write lgStrSQL
'Response.End    
End Sub

'============================================================================================================
' Name : SubCreatData
' Desc : Query Data from Db
'============================================================================================================
Sub SubCreatData()
    Dim iLoopMax
    Dim iKey1
    Dim lgStrSQL
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
 
 	lgStrSQL = " DELETE HFA150T "
	lgStrSQL = lgStrSQL & " WHERE emp_no =" & FilterVar(lgKeyStream(0), "''", "S") & " and year_yy = " &  FilterVar(lgKeyStream(2), "''", "S")

'Response.Write lgStrSQL
'Response.End 
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
 
	call CommonQueryRs(" count(*) "," HFA150T "," YEAR_YY =" &  FilterVar(lgKeyStream(2)-1, "''", "S") & " and EMP_NO = " & FilterVar(lgKeyStream(0), "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

' INSERT 

    lgStrSQL = "INSERT INTO HFA150T(YEAR_YY, EMP_NO, FAMILY_NAME, FAMILY_REL, FAMILY_RES_NO, "
    lgStrSQL = lgStrSQL & " BASE_YN, PARIA_YN, CHILD_YN, INSUR_YN, MEDI_YN, EDU_YN, CARD_YN, NAT_FLAG, ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT) "	
    
  	'If Replace(lgF0, Chr(11), "")  > 0 Then
  	if 1=2 then
 
		lgStrSQL = lgStrSQL & "SELECT " &  FilterVar(lgKeyStream(2), "''", "S") & "," & FilterVar(lgKeyStream(0), "''", "S")  
		lgStrSQL = lgStrSQL & "		,FAMILY_NAME, FAMILY_REL, FAMILY_RES_NO, BASE_YN, PARIA_YN, CHILD_YN, INSUR_YN, MEDI_YN, EDU_YN, CARD_YN, NAT_FLAG, "
		lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & "," & FilterVar(lgSvrDateTime, "''", "S") & "," & FilterVar(lgKeyStream(0), "''", "S") & "," & FilterVar(lgSvrDateTime, "''", "S") 
		lgStrSQL = lgStrSQL & " FROM  HFA150T "
		lgStrSQL = lgStrSQL & " WHERE YEAR_YY =" &  FilterVar(lgKeyStream(2)-1, "''", "S") & " and EMP_NO = " & FilterVar(lgKeyStream(0), "''", "S")
  		
	Else
		lgStrSQL = lgStrSQL &" SELECT " &  FilterVar(lgKeyStream(2), "''", "S") & "," & FilterVar(lgKeyStream(0), "''", "S") 
		lgStrSQL = lgStrSQL & "		, family_nm FAMILY_NAME,  AA.REFERENCE FAMILY_REL,  res_no FAMILY_RES_NO "
		lgStrSQL = lgStrSQL & "		,'N' BASE_YN,'N' PARIA_YN,'N' CHILD_YN,'N' INSUR_YN,'N' MEDI_YN,'N' EDU_YN,'N' CARD_YN, '1' ,"
		lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & "," & FilterVar(lgSvrDateTime, "''", "S") & "," & FilterVar(lgKeyStream(0), "''", "S") & "," & FilterVar(lgSvrDateTime, "''", "S") 
	
		lgStrSQL = lgStrSQL & " FROM  HAA020T  H LEFT JOIN "
		lgStrSQL = lgStrSQL & " 	( SELECT A.MINOR_CD ,A.MINOR_NM,B.REFERENCE "
		lgStrSQL = lgStrSQL & " 	  FROM  B_MINOR A  JOIN B_CONFIGURATION B  ON  A.MAJOR_CD=B.MAJOR_CD AND A.MINOR_CD=B.MINOR_CD "
		lgStrSQL = lgStrSQL & " 	  WHERE A.MAJOR_CD='H0023' ) AA 	ON  H.rel_cd = AA.MINOR_CD  "
		lgStrSQL = lgStrSQL & " WHERE emp_no =" & FilterVar(lgKeyStream(0), "''", "S")  
 	End If 
'Response.Write lgStrSQL
'Response.End
     
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)  
 
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
       Case "UID_M0003"  
            With Parent
               .DBQuery(1)
	        End with                  
    End Select    
       
</Script>	
