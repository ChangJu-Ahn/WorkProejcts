<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	lgSvrDateTime = GetSvrDateTime
    Call HideStatusWnd_uniSIMS

                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        
    'Multi SpreadSheet

'    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
'    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '☜: Save,Update
             Call SubBizSaveSingleUpdate()
'             Call SubBizSaveMulti()
'        Case CStr(UID_M0003)                                                         '☜: Delete
'             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1


    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
        lgErrorStatus = "YES"
        if  lgPrevNext = "N" then
            Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)
            lgPrevNext = ""
            Call SubBizQuery()
        elseif lgPrevNext = "P" then
            Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)
            lgPrevNext = ""
            Call SubBizQuery()
        end if
        exit sub
    end if 


    'iKey1 = FilterVar(lgKeyStream(0),"''", "S")     ' 사번으로조회 
    'iKey1 = iKey1 & " AND internal_cd LIKE '" & lgKeyStream(1) & "%'"

    Call SubMakeSQLStatements("R",emp_no)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
%>
<Script Language=vbscript>
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	' Set condition area, contents area
	'--------------------------------------------------------------------------------------------------------
       With Parent	

            .frm1.txtpay_cd.value = "<%=ConvSPChars(lgObjRs("pay_cd"))%>"   
                 
            .frm1.txtAnnualSal.value = "<%=UNINumClientFormat(lgObjRs("annual_sal"), ggAmtOfMoney.DecPoint, 0)%>"
            .frm1.txtSalary.value = "<%=UNINumClientFormat(lgObjRs("salary"), ggAmtOfMoney.DecPoint, 0)%>"
            .frm1.txtBonusSalary.value = "<%=UNINumClientFormat(lgObjRs("bonus_salary"), ggAmtOfMoney.DecPoint, 0)%>"   
            
            .frm1.txtBankNm.value = "<%=FuncCodeName(6, "", ConvSPChars(lgObjRs("bank")))%>"
            .frm1.txtAccntNo.value = "<%=lgObjRs("bank_accnt")%>"

            If "<%=ConvSPChars(lgObjRs("trade_union"))%>" = "Y" Then   '노조 
                .frm1.rdoUnionFlag1.value = "Y"
                .frm1.rdoUnionFlag1.checked = true
                .frm1.rdoUnionFlag2.checked = false
            Else
                .frm1.rdoUnionFlag2.value = "N"
                .frm1.rdoUnionFlag2.checked = true
                .frm1.rdoUnionFlag1.checked = false
            End If

            If "<%=ConvSPChars(lgObjRs("press_gubun"))%>" = "Y" Then   '기자 
                .frm1.rdoPressFlag1.value = "Y"
                .frm1.rdoPressFlag1.checked = true
                .frm1.rdoPressFlag2.checked = false
            Else
                .frm1.rdoPressFlag2.value = "N"
                .frm1.rdoPressFlag2.checked = true
                .frm1.rdoPressFlag1.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("oversea_labor_gubun"))%>" = "Y" Then   '국내거주 
                .frm1.rdoOverseaFlag1.value = "Y"
                .frm1.rdoOverseaFlag1.checked = true
                .frm1.rdoOverseaFlag2.checked = false
            Else
                .frm1.rdoOverseaFlag2.value = "N"
                .frm1.rdoOverseaFlag2.checked = true
                .frm1.rdoOverseaFlag1.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("res_flag"))%>" = "Y" Then   '거주구분 
                .frm1.rdoResFlag1.value = "Y"
                .frm1.rdoResFlag1.checked = true
                .frm1.rdoResFlag2.checked = false
            Else
                .frm1.rdoResFlag2.value = "N"
                .frm1.rdoResFlag2.checked = true
                .frm1.rdoResFlag1.checked = false
            End If
                        
            If "<%=ConvSPChars(lgObjRs("prov_type"))%>" = "Y" Then
                .frm1.chkPayFlg.checked = true
            Else
                .frm1.chkPayFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("employ_insur"))%>" = "Y" Then
                .frm1.chkEmpInsurFlg.checked = true
            Else
                .frm1.chkEmpInsurFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("year_calcu"))%>" = "Y" Then
                .frm1.chkYearFlg.checked = true
            Else
                .frm1.chkYearFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("retire_give"))%>" = "Y" Then
                .frm1.chkRetireFlg.checked = true
            Else
                .frm1.chkRetireFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("tax_calcu"))%>" = "Y" Then
                .frm1.chkTaxFlg.checked = true
            Else
                .frm1.chkTaxFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("year_mon_give"))%>" = "Y" Then
                .frm1.chkYearTaxFlg.checked = true
            Else
                .frm1.chkYearTaxFlg.checked = false
            End If
                        
            If "<%=ConvSPChars(lgObjRs("spouse"))%>" = "Y" Then
                .frm1.chkSpouseFlg.checked = true
            Else
                .frm1.chkSpouseFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("lady"))%>" = "Y" Then
                .frm1.chkLadyFlg.checked = true
            Else
                .frm1.chkLadyFlg.checked = false
            End If

            .frm1.txtChild.value = "<%=ConvSPChars(lgObjRs("chl_rear"))%>"
            .frm1.txtOld.value = "<%=ConvSPChars(lgObjRs("supp_old_cnt"))%>"
            .frm1.txtYoung.value = "<%=ConvSPChars(lgObjRs("supp_young_cnt"))%>"
            .frm1.txtParia.value = "<%=ConvSPChars(lgObjRs("paria_cnt"))%>"
            .frm1.txtOldCnt.value = "<%=ConvSPChars(lgObjRs("old_cnt"))%>"
            .frm1.txttax_cd.value = "<%=ConvSPChars(lgObjRs("tax_cd"))%>"              

       End With          
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
</Script>       
<%
    End If
    Call SubCloseRs(lgObjRs)
    
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
End Sub    

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  HAA010T"
    lgStrSQL = lgStrSQL & " SET " 
    '병역구분 char(2)
    lgStrSQL = lgStrSQL & " mil_type = " & FilterVar(UCase(Request("txtmil_type")), "''", "S") & ","
    '병역군별 char(2)
    lgStrSQL = lgStrSQL & " mil_kind = " & FilterVar(UCase(Request("txtmil_kind")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " mil_start = " & FilterVar(Request("txtmil_start"),"NULL","S") & ","    ' datetime
    lgStrSQL = lgStrSQL & " mil_end = " & FilterVar(Request("txtmil_end"),"NULL","S") & ","        ' datetime
    '병역등급 char(2)
    lgStrSQL = lgStrSQL & " mil_grade = " & FilterVar(UCase(Request("txtmil_grade")), "''", "S") & ","
    '병역병과 char(2)
    lgStrSQL = lgStrSQL & " mil_branch = " & FilterVar(UCase(Request("txtmil_branch")), "''", "S") & ","

    lgStrSQL = lgStrSQL & " updt_emp_no = " & FilterVar(gUsrId, "''", "S") & ","                                   ' char(10)
    lgStrSQL = lgStrSQL & " updt_dt = " & FilterVar(lgSvrDateTime, "''", "S") & ","                ' datetime
    '군번 char(10)
    lgStrSQL = lgStrSQL & " mil_no = " & FilterVar(UCase(Request("txtmil_no")), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no =  " & FilterVar(Request("txtEmp_no"), "''", "S") & ""

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
                Case ""
                      lgStrSQL = "Select " 
                      lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ",pay_cd)  as pay_cd,"                      
	                  lgStrSQL = lgStrSQL & " annual_sal ,  salary ,"
	                  lgStrSQL = lgStrSQL & " bonus_salary ,"
	                  lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0006", "''", "S") & ",tax_cd)  as tax_cd,"
	                  lgStrSQL = lgStrSQL & " bank , bank_accnt  ,  chl_rear , supp_old_cnt ,"
	                  lgStrSQL = lgStrSQL & " supp_young_cnt  ,  paria_cnt  ,  old_cnt ,"
	                  lgStrSQL = lgStrSQL & " trade_union ,  press_gubun , oversea_labor_gubun ,spouse , lady , "
	                  lgStrSQL = lgStrSQL & " res_flag ,prov_type  ,  employ_insur ,  year_calcu  ,  retire_give  ,  tax_calcu  ,  year_mon_give " 
                      lgStrSQL = lgStrSQL & " From  HDF020T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode 	
                Case "P"
                      lgStrSQL = "Select TOP 1 * " 
                      lgStrSQL = lgStrSQL & " From  HDF020T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 	
                      lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
                Case "N"
                      lgStrSQL = "Select TOP 1 * " 
                      lgStrSQL = lgStrSQL & " From  HDF020T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
                      lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC "
             End Select
      Case "C"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
      Case "U"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
      Case "D"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
    End Select
	'------ Developer Coding part (End   ) ------------------------------------------------------------------


    lgStrSQL = "Select " 
    lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ",pay_cd)  as pay_cd,"                      
	lgStrSQL = lgStrSQL & " annual_sal ,  salary ,"
	lgStrSQL = lgStrSQL & " bonus_salary ,"
    lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0006", "''", "S") & ",tax_cd)  as tax_cd,"
	lgStrSQL = lgStrSQL & " bank , bank_accnt  ,  chl_rear , supp_old_cnt ,"
	lgStrSQL = lgStrSQL & " supp_young_cnt  ,  paria_cnt  ,  old_cnt ,"
	lgStrSQL = lgStrSQL & " trade_union ,  press_gubun , oversea_labor_gubun ,spouse , lady , "
	lgStrSQL = lgStrSQL & " res_flag ,prov_type  ,  employ_insur ,  year_calcu  ,  retire_give  ,  tax_calcu  ,  year_mon_give " 
    lgStrSQL = lgStrSQL & " From  HDF020T "
    lgStrSQL = lgStrSQL & " WHERE emp_no =  " & FilterVar(pCode , "''", "S") & ""

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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)    'Can not create(Demo code)
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
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
        
