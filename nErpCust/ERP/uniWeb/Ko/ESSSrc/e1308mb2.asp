<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->

<%
	Dim txtMainInsertFlag
	lgSvrDateTime = GetSvrDateTime
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd_uniSIMS

    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")    
    txtMainInsertFlag        = Request("strMainInsertFlag")            

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
   Dim iKey1,iRet,strWhere

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iRet = SubEmpBase1(lgKeyStream(0),lgKeyStream(13),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
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
		strWhere = " emp_no=" & FilterVar(lgKeyStream(0),"''", "S")
		strWhere = strWhere & " AND retire_dt is null"       
		Call CommonQueryRs(" internal_cd "," HAA010T ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
		if gProAuth = 0 then
			if lgF0="X" or lgF0="" then
	%>
	        <Script Language=vbscript>
	            With parent.parent
	                .txtEmp_no2.Value = "<%=lgKeyStream(0)%>"
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
				Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)
				Response.end		
			end if	
		else				
			if inStr(ConvSPChars(lgF0),ConvSPChars(lgKeyStream(1)))=0 then

	%>
	        <Script Language=vbscript>
	            With parent.parent
	                '.txtEmp_no2.Value = "<%=lgKeyStream(0)%>"
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
				if lgF0="X" or lgF0="" then
					Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)
				else 
					Call DisplayMsgBox("800454", vbInformation, "", "", I_MKSCRIPT)
				end if
				Response.end
			end if    
        end if    
        if  lgPrevNext = "N" then
            Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)
        elseif lgPrevNext = "P" then
            Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)
        end if
        Response.End
    End If

    If txtMainInsertFlag =  "Y" Then
        Response.End
    End If    
'------------  
    
	 iKey1 = FilterVar(lgKeyStream(0),"''", "S")
	 iKey1 = iKey1 &" and YEAR_YY = "  &  FilterVar(lgKeyStream(1),"''", "S")	     
	 iKey1 = iKey1 &" and FAMILY_RES_NO = "  &  FilterVar(lgKeyStream(4),"''", "S")
	 
    Call SubMakeSQLStatements("R",iKey1)  
                                         '☜ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
%>

<Script Language=vbscript>
       With Parent	
			.Frm1.txtYear.Value			= "<%=ConvSPChars(lgKeyStream(1))%>"                 
            .frm1.txtFamilyName.value	= "<%=ConvSPChars(lgObjRs("FAMILY_NAME"))%>"         
            .frm1.txtRel_cd.value		= "<%=ConvSPChars(lgObjRs("FAMILY_REL"))%>"        
            .frm1.txtRes_no.value		= "<%=ConvSPChars(lgObjRs("FAMILY_RES_NO"))%>"
            .frm1.txtBase_YN.value		= "<%=ConvSPChars(lgObjRs("BASE_YN"))%>"
            .frm1.txtParia_YN.value		= "<%=ConvSPChars(lgObjRs("PARIA_YN"))%>"
            .frm1.txtChild_YN.value		= "<%=ConvSPChars(lgObjRs("CHILD_YN"))%>"
            .frm1.txtInsur_YN.value		= "<%=ConvSPChars(lgObjRs("INSUR_YN"))%>"
            .frm1.txtMedi_YN.value		= "<%=ConvSPChars(lgObjRs("MEDI_YN"))%>"
            .frm1.txtEdu_YN.value		= "<%=ConvSPChars(lgObjRs("EDU_YN"))%>"
            .frm1.txtCard_YN.value		= "<%=ConvSPChars(lgObjRs("CARD_YN"))%>"
            .frm1.txtNat_flag.value		= "<%=ConvSPChars(lgObjRs("NAT_FLAG"))%>"
       End With          

</Script>       
<%   
	End if
	
    Call SubCloseRs(lgObjRs)
    	
End Sub     

'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
	Dim counts
	Dim i
	Dim strInput_emp_no
	Dim strClose_type
	Dim strClose_dt
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
 
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
			Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                             '☜ : Update
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

    lgStrSQL = "DELETE  HFA150T"
    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & "         EMP_NO		= " & FilterVar(lgKeyStream(0),"''","S")
    lgStrSQL = lgStrSQL & "   AND   YEAR_YY		= " & FilterVar(lgKeyStream(1),"''","S")
    lgStrSQL = lgStrSQL & "   AND   FAMILY_RES_NO = " & FilterVar(lgKeyStream(4),"''","S")
    
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords


'Response.Write lgStrSQL
'Response.End    
    End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	'중복 데이터를 check 한다 
	Call CommonQueryRs(" count(EMP_NO) "," HFA150T "," EMP_NO = " & FilterVar(lgKeyStream(0),"''","S")& " AND YEAR_YY = " & FilterVar(lgKeyStream(1),"''","S") &_
	"   AND FAMILY_RES_NO = "	& FilterVar(lgKeyStream(4),"''","S")  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Trim(Replace(lgF0,Chr(11),"")) = 0 then
	Else
	    Call DisplayMsgBox("800446", vbInformation, "", "", I_MKSCRIPT)    
        lgErrorStatus = "YES"
	    Exit sub                                   
	End if
			
    lgStrSQL = "INSERT INTO HFA150T(YEAR_YY, EMP_NO, FAMILY_NAME, FAMILY_REL, FAMILY_RES_NO, "
    lgStrSQL = lgStrSQL & " BASE_YN, PARIA_YN, CHILD_YN, INSUR_YN, MEDI_YN, EDU_YN, CARD_YN,NAT_FLAG, ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT) "
	lgStrSQL = lgStrSQL & " VALUES(" 
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(2), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(4), "''", "S")   & ","	
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(5), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(6), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(7), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(8), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(9), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(10), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(11), "''", "S")   & ","	
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(12), "''", "S")   & ","	
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")			& "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")			& "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") 
	lgStrSQL = lgStrSQL & ")"
 
 
'Response.Write lgStrSQL
'Response.End
 
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
'Response.Write "ERR:" & ERR.Description
'Response.End    
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
	
End Sub


'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgStrSQL = "UPDATE  HFA150T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " BASE_YN			= " & FilterVar(lgKeyStream(5),"''","S") & ","
    lgStrSQL = lgStrSQL & " PARIA_YN		= " & FilterVar(lgKeyStream(6),"''","S") & ","
    lgStrSQL = lgStrSQL & " CHILD_YN		= " & FilterVar(lgKeyStream(7),"''","S") & ","
    lgStrSQL = lgStrSQL & " INSUR_YN		= " & FilterVar(lgKeyStream(8),"''","S") & ","
    lgStrSQL = lgStrSQL & " MEDI_YN			= " & FilterVar(lgKeyStream(9),"''","S") & ","
    lgStrSQL = lgStrSQL & " EDU_YN			= " & FilterVar(lgKeyStream(10),"''","S") & ","
    lgStrSQL = lgStrSQL & " CARD_YN			= " & FilterVar(lgKeyStream(11),"''","S") & "," 
    lgStrSQL = lgStrSQL & " NAT_FLAG		= " & FilterVar(lgKeyStream(12),"''","S") & "," 
    
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = "	& FilterVar(lgKeyStream(0), "''", "S")	 & ","
    lgStrSQL = lgStrSQL & " UPDT_DT = "		&  FilterVar(lgSvrDateTime, "''", "S") 
    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & "         EMP_NO		= " & FilterVar(lgKeyStream(0),"''","S")
    lgStrSQL = lgStrSQL & "   AND   YEAR_YY		= " & FilterVar(lgKeyStream(1),"''","S")
    lgStrSQL = lgStrSQL & "   AND   FAMILY_RES_NO = " & FilterVar(lgKeyStream(4),"''","S")

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

            lgStrSQL = "Select FAMILY_NAME, FAMILY_REL , dbo.ufn_GetCodeName('H0140',FAMILY_REL) as FAMILY_REL_NM, "
            lgStrSQL = lgStrSQL & "  FAMILY_RES_NO , BASE_YN, PARIA_YN, CHILD_YN, INSUR_YN, MEDI_YN, EDU_YN, CARD_YN,NAT_FLAG"
            lgStrSQL = lgStrSQL & " From HFA150T"
            lgStrSQL = lgStrSQL & " WHERE EMP_NO = " & pCode
            lgStrSQL = lgStrSQL & " ORDER BY FAMILY_REL, FAMILY_RES_NO ASC"                	
           
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
       Case "UID_M0003"

          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    

</Script>	
