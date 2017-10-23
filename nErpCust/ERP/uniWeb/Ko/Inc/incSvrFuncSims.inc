<%   
'==============================================================================
'==============================================================================
' uniSIMS related function(Server Side)
'==============================================================================
'==============================================================================

'==============================================================================
' Hide Current Window
'==============================================================================
Sub HideStatusWnd_uniSIMS()
	Response.Write "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
	Response.Write "Sub Document_onReadyStateChange()" & vbCrLf
	Response.Write " On Error Resume Next "            & vbCrLf
	Response.Write "Call parent.BtnDisabled(False)"    & vbCrLf	
	Response.Write "Call parent.LayerShowHide(0)"      & vbCrLf
	Response.Write "End Sub"  & vbCrLf
	Response.Write "</" & "Script" & ">" & vbCrLf
End Sub

Sub SubEmpBase(Emp_no,Internal_cd,Con,oEmp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)

    Dim lgStrSQL
    Dim iObjConn
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    Call SubOpenDB(iObjConn)    

    oEmp_no = ""    
    Name = ""
    roll_pstn = ""
    dept_nm = ""
    resent_promote_dt = ""
    group_entr_dt = ""
    entr_dt = ""
    
    lgStrSQL = "Select top 1 emp_no, name, dept_nm, entr_dt, group_entr_dt, resent_promote_dt, " 
    lgStrSQL = lgStrSQL & " (select b_minor.minor_nm from b_minor where b_minor.minor_cd = roll_pstn and b_minor.major_cd='H0002') as roll_pstn"
    lgStrSQL = lgStrSQL & " From  HAA010T "
    Select Case Con
        Case "N"
            lgStrSQL = lgStrSQL & " WHERE emp_no > " & FilterVar(emp_no,"''", "S")
            lgStrSQL = lgStrSQL & "   AND internal_cd LIKE '" & Internal_cd & "%' ORDER BY emp_no ASC"
        Case "P"
            lgStrSQL = lgStrSQL & " WHERE emp_no < " & FilterVar(emp_no,"''", "S")
            lgStrSQL = lgStrSQL & "   AND internal_cd LIKE '" & Internal_cd & "%' ORDER BY emp_no DESC"
        Case Else
            lgStrSQL = lgStrSQL & " WHERE emp_no = '" & emp_no & "'"
	End Select
    If 	FncOpenRs("R",iObjConn,iObjRs,lgStrSQL,"X","X") = False Then
    Else
        oEmp_no = iObjRs("emp_no")
        Name = iObjRs("name")
        roll_pstn = iObjRs("roll_pstn")
        dept_nm = iObjRs("dept_nm")
        resent_promote_dt = iObjRs("resent_promote_dt")
        group_entr_dt = iObjRs("group_entr_dt")
        entr_dt = iObjRs("entr_dt")
    End If
    
    Call SubCloseRs(iObjRs)
    Call SubCloseDB(iObjConn)        

End Sub

Function SubEmpBase1(Emp_no,Internal_cd,Con,oEmp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)

    Dim lgStrSQL
    Dim iObjConn
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    Call SubOpenDB(iObjConn)    
    SubEmpBase1 = True
    oEmp_no = ""    
    Name = ""
    roll_pstn = ""
    dept_nm = ""
    resent_promote_dt = ""
    group_entr_dt = ""
    entr_dt = ""
    
    lgStrSQL = "Select top 1 emp_no, name, dept_nm, entr_dt, group_entr_dt, resent_promote_dt, " 
    lgStrSQL = lgStrSQL & " (select b_minor.minor_nm from b_minor where b_minor.minor_cd = roll_pstn and b_minor.major_cd='H0002') as roll_pstn"
    lgStrSQL = lgStrSQL & " From  HAA010T "
    Select Case Con
        Case "N"
            lgStrSQL = lgStrSQL & " WHERE emp_no > " & FilterVar(emp_no,"''", "S")
        Case "P"
            lgStrSQL = lgStrSQL & " WHERE emp_no < " & FilterVar(emp_no,"''", "S")
        Case Else
            lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(emp_no,"''", "S")
	End Select
	
	lgStrSQL = lgStrSQL & " AND retire_dt is null "  
	if gProAuth <> 0 then 
          lgStrSQL = lgStrSQL & "   AND internal_cd LIKE '" & Internal_cd & "%'  and  retire_dt is null"
	end if
	if Con="P" then
		lgStrSQL = lgStrSQL & " order by emp_no desc "
	end if

    If 	FncOpenRs("R",iObjConn,iObjRs,lgStrSQL,"X","X") = False Then
        SubEmpBase1 = False
    Else
        oEmp_no = iObjRs("emp_no")
        Name = iObjRs("name")
        roll_pstn = iObjRs("roll_pstn")
        dept_nm = iObjRs("dept_nm")
        resent_promote_dt = iObjRs("resent_promote_dt")
        group_entr_dt = iObjRs("group_entr_dt")
        entr_dt = iObjRs("entr_dt")
    End If
    
    Call SubCloseRs(iObjRs)
    Call SubCloseDB(iObjConn)        

End Function
Function SubEmpBaseDiligAuth(Emp_no,Internal_cd,DiligAuth,Con,oEmp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)

    Dim lgStrSQL
    Dim iObjConn
    Dim strDiligAuth,i

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
  
    Call SubOpenDB(iObjConn)    
    SubEmpBaseDiligAuth = True
    oEmp_no = ""    
    Name = ""
    roll_pstn = ""
    dept_nm = ""
    resent_promote_dt = ""
    group_entr_dt = ""
    entr_dt = ""

    strDiligAuth = ""
    If  Replace(DiligAuth,chr(11),"")="" Then
        strDiligAuth = "'%'"
    Else
        DiligAuth = Split(DiligAuth,chr(11))
        strDiligAuth = FilterVar(DiligAuth(0),"''", "S")
        For i=1 To UBound(DiligAuth)-1
            strDiligAuth = strDiligAuth & " or dept_cd =" & FilterVar(DiligAuth(i),"''", "S")
        Next
    End If
     

    lgStrSQL = "Select top 1 emp_no, name, dept_nm, entr_dt, group_entr_dt, resent_promote_dt, " 
    lgStrSQL = lgStrSQL & " (select b_minor.minor_nm from b_minor where b_minor.minor_cd = roll_pstn and b_minor.major_cd='H0002') as roll_pstn"
    lgStrSQL = lgStrSQL & " From  HAA010T "
    Select Case Con
        Case "N"
            lgStrSQL = lgStrSQL & " WHERE emp_no > " & FilterVar(emp_no,"''", "S")
'           lgStrSQL = lgStrSQL & "   AND internal_cd LIKE '" & Internal_cd & "%'"
            lgStrSQL = lgStrSQL & "   AND (dept_cd = "& strDiligAuth & ")"
            lgStrSQL = lgStrSQL & "   ORDER BY emp_no ASC"
            
        Case "P"
            lgStrSQL = lgStrSQL & " WHERE emp_no < " & FilterVar(emp_no,"''", "S")
'           lgStrSQL = lgStrSQL & "   AND internal_cd LIKE '" & Internal_cd & "%'"
            lgStrSQL = lgStrSQL & "   AND (dept_cd = "& strDiligAuth & ")"
            lgStrSQL = lgStrSQL & "   ORDER BY emp_no DESC"
        Case Else
            lgStrSQL = lgStrSQL & " WHERE emp_no = '" & emp_no & "'"
	End Select

    If 	FncOpenRs("R",iObjConn,iObjRs,lgStrSQL,"X","X") = False Then
        SubEmpBaseDiligAuth = False
    Else
  
        oEmp_no = iObjRs("emp_no")
'        emp_no = iObjRs("emp_no")

        Name = iObjRs("name")
        roll_pstn = iObjRs("roll_pstn")
        dept_nm = iObjRs("dept_nm")
        resent_promote_dt = iObjRs("resent_promote_dt")
        group_entr_dt = iObjRs("group_entr_dt")
        entr_dt = iObjRs("entr_dt")
    End If
    
    Call SubCloseRs(iObjRs)
    Call SubCloseDB(iObjConn)        

End Function
Function EmpBaseDiligAuthCheck(Emp_no,oDiligAuth,oAuthCheck,login_emp_no,Con,oEmp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)

    Dim lgStrSQL
    Dim iObjConn
    Dim strDiligAuth,i
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
  
    Call SubOpenDB(iObjConn)    
    EmpBaseDiligAuthCheck = True
    oEmp_no = ""    
    Name = ""
    roll_pstn = ""
    dept_nm = ""
    resent_promote_dt = ""
    group_entr_dt = ""
    entr_dt = ""

	If  Replace(oDiligAuth,chr(11),"")="" Then
        strDiligAuth = " internal_cd like '%'"
    Else
		strDiligAuth =                " internal_cd in "
		strDiligAuth = strDiligAuth & " (select internal_cd from haa010t "
		strDiligAuth = strDiligAuth & "  where internal_cd in "
		strDiligAuth = strDiligAuth & "         (select a.internal_cd "
		strDiligAuth = strDiligAuth & "             from haa010t a, e11090t b "
		strDiligAuth = strDiligAuth & "            where a.internal_cd like  rtrim(b.internal_cd) + '%' AND"
		strDiligAuth = strDiligAuth & "                  b.internal_auth = 'Y' AND "
		strDiligAuth = strDiligAuth & "                  b.emp_no = '" & login_emp_no & "') "
		strDiligAuth = strDiligAuth & "        and internal_cd not in "
		strDiligAuth = strDiligAuth & "         (select a.internal_cd "
		strDiligAuth = strDiligAuth & "            from haa010t a, e11090t b "
		strDiligAuth = strDiligAuth & "           where a.internal_cd like  rtrim(b.internal_cd) + '%' and "
		strDiligAuth = strDiligAuth & "                 b.emp_no =  '" & login_emp_no & "' and b.internal_auth = 'N') "
		strDiligAuth = strDiligAuth & " union "
		strDiligAuth = strDiligAuth & " (select internal_cd from e11090t "
		strDiligAuth = strDiligAuth & " where emp_no = '" & login_emp_no & "' and internal_auth = 'N')) "
	end if
	
	lgStrSQL = "Select  top 1 emp_no,name,dept_nm, res_no, dbo.ufn_GetCodeName('H0002',roll_pstn) as roll_pstn " 
	lgStrSQL = lgStrSQL & " From  HAA010T "
	
    Select Case Con
        Case "N"
            lgStrSQL = lgStrSQL & " WHERE emp_no >" &FilterVar(lgKeyStream(0),"''", "S")  
            lgStrSQL = lgStrSQL & " and  ( "& strDiligAuth &"  or emp_no=" & FilterVar(login_emp_no,"''", "S")& " ) "
            lgStrSQL = lgStrSQL & "  AND   retire_dt is null ORDER BY emp_no ASC"
           
        Case "P"
            lgStrSQL = lgStrSQL & " WHERE emp_no <" &FilterVar(lgKeyStream(0),"''", "S")  
            lgStrSQL = lgStrSQL & " and  ( "& strDiligAuth &"  or emp_no=" & FilterVar(login_emp_no,"''", "S")& " ) "
            lgStrSQL = lgStrSQL & "  AND   retire_dt is null ORDER BY emp_no DESC"

        Case Else
            lgStrSQL = lgStrSQL & " WHERE emp_no =" &FilterVar(lgKeyStream(0),"''", "S")  
            lgStrSQL = lgStrSQL & " and  ( "& strDiligAuth &"  or emp_no=" & FilterVar(login_emp_no,"''", "S")& " ) "
            lgStrSQL = lgStrSQL & "  AND   retire_dt is null "
	End Select
	

    If 	FncOpenRs("R",iObjConn,iObjRs,lgStrSQL,"X","X") = False Then
        EmpBaseDiligAuthCheck = False
    Else
  
        oEmp_no           = iObjRs("emp_no")
        Name              = iObjRs("name")
        roll_pstn         = iObjRs("roll_pstn")
        dept_nm           = iObjRs("dept_nm")
        resent_promote_dt = iObjRs("resent_promote_dt")
        group_entr_dt     = iObjRs("group_entr_dt")
        entr_dt           = iObjRs("entr_dt")
    End If
    
    Call SubCloseRs(iObjRs)
    Call SubCloseDB(iObjConn)        

End Function
%>
<Script Language=VBScript >
    Dim gServerIP
    Dim gLogoName
    
    If Instr(document.location.href,"http://") > 0 Then
       gServerIP = "http://<%= request.servervariables("server_name") %>"
    Else
       gServerIP = "https://<%= request.servervariables("server_name") %>"
    End If   

    gLogoName = "<%=Request.Cookies("unierp")("gLogoName")%>"
</Script>
