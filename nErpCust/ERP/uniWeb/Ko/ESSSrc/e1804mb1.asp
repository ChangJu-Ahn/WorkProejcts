<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
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

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '☜: Save,Update
             Call SubBizSaveMulti()
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

    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
        lgstrData = ""
        iDx       = 1
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lang_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("menu_id"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("menu_name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("href"))
            if ConvSPChars(lgObjRs("menu_level")) = "1" then
                lgstrData = lgstrData & Chr(11) & "메뉴"
            elseif  ConvSPChars(lgObjRs("menu_level")) = "2" then
                lgstrData = lgstrData & Chr(11) & "프로그램"
            else
                lgstrData = lgstrData & Chr(11) & ""
            end if
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pro_use_flag"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pro_auth"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ref_menu_id"))

            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 
    End If
    Call SubCloseRs(lgObjRs)

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
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  E11070T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " app_yn = " & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Updt_emp_no = " & FilterVar(Request("txtUpdtUserId"), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_strt_dt = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = " & FilterVar(UCase(arrColVal(4)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

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
                    lgStrSQL = "Select lang_cd,menu_id, menu_name, href, menu_level, pro_use_flag, ref_menu_id, pro_type, "
                    lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0120", "''", "S") & ",pro_auth) as pro_auth "
                    lgStrSQL = lgStrSQL & " From E11000T"
                    lgStrSQL = lgStrSQL & " WHERE lang_cd = " & FilterVar(lgKeyStream(2), "''", "S") & " and menu_level in (1,2)" 
                    lgStrSQL = lgStrSQL & "  ORDER BY ref_menu_id,orders"
                Case "P"
                    lgStrSQL = "Select emp_no, "
                    lgStrSQL = lgStrSQL & " dilig_dt, dilig_cd,"
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetCodeName(" & FilterVar("hca010t", "''", "S") & ",dilig_cd,'') as dilig_nm , "
                    lgStrSQL = lgStrSQL & " dilig_cnt, "
                    lgStrSQL = lgStrSQL & " dilig_hh, "
                    lgStrSQL = lgStrSQL & " dilig_mm, "
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(app_emp_no)  as app_emp_no, "
                    lgStrSQL = lgStrSQL & " app_yn "
                    lgStrSQL = lgStrSQL & " From E11060T"
                    lgStrSQL = lgStrSQL & " WHERE app_emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & " Order by dilig_dt DESC"
                    'lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode
                Case "N"
                    lgStrSQL = "Select emp_no, "
                    lgStrSQL = lgStrSQL & " dilig_dt, dilig_cd,"
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetCodeName(" & FilterVar("hca010t", "''", "S") & ",dilig_cd,'') as dilig_nm , "
                    lgStrSQL = lgStrSQL & " dilig_cnt, "
                    lgStrSQL = lgStrSQL & " dilig_hh, "
                    lgStrSQL = lgStrSQL & " dilig_mm, "
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(app_emp_no)  as app_emp_no, "
                    lgStrSQL = lgStrSQL & " app_yn "
                    lgStrSQL = lgStrSQL & " From E11060T"
                    lgStrSQL = lgStrSQL & " WHERE app_emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
                    lgStrSQL = lgStrSQL & " Order by dilig_dt DESC"
                    'lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
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
          Else
             'parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	
