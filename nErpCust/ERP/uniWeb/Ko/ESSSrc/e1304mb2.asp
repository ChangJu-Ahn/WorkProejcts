<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<%

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd_uniSIMS
                                                               '��: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        
    'Multi SpreadSheet

'    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '��: Fetch count at a time for VspdData
'    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '��: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '��: Save,Update
             Call SubBizSave()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim strEmpNo  
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    strEmpNo  = lgKeyStream(0)

    Call SubEmpBase(lgKeyStream(0),lgKeyStream(1),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
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
    if emp_no = "" then
        return
    end if 

    strEmpNo  = emp_no
    
    Call SubMakeSQLStatements("R",strEmpNo)                                       '�� : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
    Else
   
%>
<Script Language=vbscript>
 
       With Parent	
            .Frm1.txtEntr_dt.Value				= "<%=UNIDateClientFormat(lgObjRs("H_ENTR_DT"))%>"
            .Frm1.txtRetire_yyyy.Value			= "<%=UNIDateClientFormat(lgObjRs("ENTR_DT"))%>"
            .Frm1.txtRetire_dt.Value			= ""
            .frm1.txtAvr_wages_amt.value        = "<%=UNINumClientFormat(lgObjRs("AVR_WAGES"), ggAmtOfMoney.DecPoint,0)%>"         
            .frm1.txtTot_duty_mm.value			= "<%=UNINumClientFormat(lgObjRs("TOT_DUTY_MM"), 0,0)%>"          
            .frm1.txtTot_prov_amt.value			= "<%=UNINumClientFormat(lgObjRs("TOT_PROV_AMT"), ggAmtOfMoney.DecPoint,0)%>"          
            .frm1.txtIncome_amt.value			= "<%=UNINumClientFormat(lgObjRs("TOT_PROV_AMT"), ggAmtOfMoney.DecPoint,0)%>"        
       End With 

</Script>       
<%     
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
                      lgStrSQL = "Select AVR_WAGES,TOT_DUTY_MM,TOT_PROV_AMT,a.entr_dt,b.entr_dt as H_ENTR_DT" 
                      lgStrSQL = lgStrSQL & " from  HGA081T AS a LEFT OUTER JOIN HAA010T AS b ON (a.emp_no=b.emp_no)"
                      lgStrSQL = lgStrSQL & " WHERE a.emp_no =  " & FilterVar(pCode , "''", "S") & ""
                Case "P"
                      lgStrSQL = "Select TOP 1 Select AVR_WAGES,TOT_DUTY_MM,TOT_PROV_AMT,a.entr_dt,b.entr_dt as H_ENTR_DT" 
                      lgStrSQL = lgStrSQL & " from  HGA081T AS a LEFT OUTER JOIN HAA010T AS b ON (a.emp_no=b.emp_no)"
                      lgStrSQL = lgStrSQL & " WHERE a.emp_no <  " & FilterVar(pCode, "''", "S") & ""
                      lgStrSQL = lgStrSQL & " ORDER BY a.emp_no DESC "
                Case "N"
                      lgStrSQL = "Select TOP 1 Select AVR_WAGES,TOT_DUTY_MM,TOT_PROV_AMT,a.entr_dt,b.entr_dt as H_ENTR_DT" 
                      lgStrSQL = lgStrSQL & " from  HGA081T AS a LEFT OUTER JOIN HAA010T AS b ON (a.emp_no=b.emp_no)"
                      lgStrSQL = lgStrSQL & " WHERE a.emp_no >  " & FilterVar(pCode, "''", "S") & ""
                      lgStrSQL = lgStrSQL & " ORDER BY a.emp_no ASC "
             End Select
    End Select
 Response.Write lgStrSQL
'Response.End
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      'Can not create(Demo code)
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      'Can not create(Demo code)
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
       Case "UID_M0001"                                                         '�� : Query
         If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                 .DBQueryOk        
	          End with
	      Else
              With Parent
                 .DBQueryFail
	          End with
          End If   
       Case "UID_M0002"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.ExeReflectOk()
          Else
            parent.ExeReflectNo()
             'Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	
