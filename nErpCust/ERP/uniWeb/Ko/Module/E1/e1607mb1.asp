<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd_uniSIMS
    
    Dim emp_no
    Dim name
    Dim dept_nm
    Dim iDx
                                                               '��: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '��: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '��: Save,Update
             Call SubBizSaveMulti()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    iKey1 = " WHERE app_emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    if lgKeyStream(2) = "Y" then
		iKey1 = iKey1 & "   AND app_yn = " & FilterVar("Y", "''", "S") & ""
    elseif lgKeyStream(2) = "R" then
        iKey1 = iKey1 & "   AND app_yn = " & FilterVar("R", "''", "S") & ""
    elseif lgKeyStream(2) = "N" then
        iKey1 = iKey1 & "   AND (app_yn <> " & FilterVar("Y", "''", "S") & " and app_yn <> " & FilterVar("R", "''", "S") & ")"
    end if

    if lgKeyStream(3) <> "" AND lgKeyStream(4) <> "" then
        iKey1 = iKey1 & "   AND dilig_strt_dt between" & FilterVar(lgKeyStream(3), "''", "S") & " and " & FilterVar(lgKeyStream(4), "''", "S")
        iKey1 = iKey1 & "   AND dilig_end_dt between" & FilterVar(lgKeyStream(3), "''", "S") & " and " & FilterVar(lgKeyStream(4), "''", "S")
    end if
    iKey1 = iKey1 & " AND dilig_cd not in (select dilig_cd from hca010t where  dilig_type=1) "

    Call SubMakeSQLStatements("R",iKey1)                                       '�� : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
%>
		<Script Language="VBScript">
			With Parent
				For i= 1 to 10
'					.document.all("SPREADCELL_APP_YN2_" & CStr(i))(0).disabled = True
'					.document.all("SPREADCELL_APP_YN2_" & CStr(i))(1).disabled = True
'					.document.all("SPREADCELL_APP_YN2_" & CStr(i))(0).checked = False
'					.document.all("SPREADCELL_APP_YN2_" & CStr(i))(1).checked = False
					.document.all("SPREADCELL_APP_YN2_" & CStr(i))(0).style.visibility = "hidden"
					.document.all("SPREADCELL_APP_YN2_" & CStr(i))(1).style.visibility = "hidden"	
				Next
			end with
		</Script>
		<%
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
        Call SetErrorStatus()
    Else
        lgstrData = ""
        iDx       = 1
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("dilig_strt_dt")) 
			lgstrData = lgstrData & Chr(11) & lgObjRs("dilig_hh")
			lgstrData = lgstrData & Chr(11) & lgObjRs("dilig_mm")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("remark"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("app_emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("app_yn"))
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
        Loop 
'Response.Write ",lgstrData:" & lgstrData
'Response.End

    End If
    Call SubCloseRs(lgObjRs)

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim strRowBak
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet
End Sub    

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
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

    On Error Resume Next                                                             '��: Protect system from crashing

    Err.Clear                                                                        '��: Clear Error status

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
            
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '��: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
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
    On Error Resume Next                                                             '��: Protect system from crashing
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  E11070T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " app_yn = " & FilterVar(UCase(arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Updt_emp_no = " & FilterVar(Request("txtUpdtUserId"), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_strt_dt = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = " & FilterVar(UCase(arrColVal(6)), "''", "S")
'Response.Write  ", update lgStrSQL:" & lgStrSQL
'Response.end   
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText
    if UCase(arrColVal(7)) = "Y" then
    '   �ܾ�table insert
        lgStrSQL = "INSERT INTO HCA060T("
        lgStrSQL = lgStrSQL & " EMP_NO,"
        lgStrSQL = lgStrSQL & " dilig_dt,"
        lgStrSQL = lgStrSQL & " dilig_cd,"
        lgStrSQL = lgStrSQL & " dilig_cnt,"
        lgStrSQL = lgStrSQL & " dilig_hh,"
        lgStrSQL = lgStrSQL & " dilig_mm,"
        lgStrSQL = lgStrSQL & " isrt_dt,"
        lgStrSQL = lgStrSQL & " isrt_emp_no,"
        lgStrSQL = lgStrSQL & " updt_dt,"        
        lgStrSQL = lgStrSQL & " updt_emp_no)"
        lgStrSQL = lgStrSQL & " VALUES("
        lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")     & ","
        lgStrSQL = lgStrSQL & FilterVar(uniConvdate(arrColVal(3)), "''", "S")   & "," 
        lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S")    & ","        
        lgStrSQL = lgStrSQL & "1,"        
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")    & ","
        lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S")    & ","
        lgStrSQL = lgStrSQL & FilterVar(GetSvrDate, "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(Request("txtUpdtUserId"), "''", "S")     & "," 
        lgStrSQL = lgStrSQL & FilterVar(GetSvrDate, "''", "S")    & ","      
        lgStrSQL = lgStrSQL & FilterVar(Request("txtUpdtUserId"), "''", "S")   
        lgStrSQL = lgStrSQL & ")"
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.end
        lgObjConn.Execute lgStrSQL,,adCmdText
    	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
    end if

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
                    lgStrSQL = "Select emp_no, "
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(emp_no)  as name, "
                    lgStrSQL = lgStrSQL & " dilig_strt_dt, dilig_hh,dilig_mm, dilig_cd,"
					lgStrSQL = lgStrSQL & "dbo.ufn_H_GetCodeName(" & FilterVar("hca010t", "''", "S") & ",dilig_cd,'') as dilig_nm , "                    
                    lgStrSQL = lgStrSQL & " remark,"
					lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(app_emp_no)  as app_emp_no, "
                    lgStrSQL = lgStrSQL & " app_yn "
                    lgStrSQL = lgStrSQL & " From E11070T"
                    lgStrSQL = lgStrSQL & pCode
                    lgStrSQL = lgStrSQL & " Order by dilig_strt_dt DESC"
                Case "P"
                Case "N"
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
'Response.Write "lgStrSQL:" & lgStrSQL
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
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
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
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
                 .grid1.SSSetData("<%=lgstrData%>")
                 .DBQueryOk()        
	          End with
	      Else
              With Parent
                 .DBQueryFail
	          End with
          End If   
       Case "UID_M0002"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
    End Select    
       
</Script>	
