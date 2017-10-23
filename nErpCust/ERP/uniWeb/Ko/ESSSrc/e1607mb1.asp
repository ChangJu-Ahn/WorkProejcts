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
    Dim iDx

    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '¢Ð: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '¢Ð: Save,Update
             Call SubBizSaveMulti()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    iKey1 = " WHERE app_emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    if lgKeyStream(2) = "Y" then
		iKey1 = iKey1 & "   AND app_yn = " & FilterVar("Y", "''", "S") & ""
    elseif lgKeyStream(2) = "R" then
        iKey1 = iKey1 & "   AND app_yn = " & FilterVar("R", "''", "S") & ""
    elseif lgKeyStream(2) = "N" then
        iKey1 = iKey1 & "   AND (app_yn <> " & FilterVar("Y", "''", "S") & " and app_yn <> " & FilterVar("R", "''", "S") & ")"
    end if

    if lgKeyStream(3) <> "" AND lgKeyStream(4) <> "" then
        iKey1 = iKey1 & "   AND convert(varchar(10),dilig_strt_dt,20) between" & FilterVar(lgKeyStream(3), "''", "S") & " and " & FilterVar(lgKeyStream(4), "''", "S")
        iKey1 = iKey1 & "   AND convert(varchar(10),dilig_end_dt,20) between" & FilterVar(lgKeyStream(3), "''", "S") & " and " & FilterVar(lgKeyStream(4), "''", "S")
    end if
    iKey1 = iKey1 & " AND dilig_cd not in (select dilig_cd from hca010t where  dilig_type=1) "

    Call SubMakeSQLStatements("R",iKey1)                                       '¢Ð : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
%>
		<Script Language="VBScript">
			With Parent
				For i= 1 to 10
					.document.all("SPREADCELL_APP_YN2_" & CStr(i))(0).style.visibility = "hidden"
					.document.all("SPREADCELL_APP_YN2_" & CStr(i))(1).style.visibility = "hidden"	
				Next
			end with
		</Script>
		<%
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
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

    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    Err.Clear                                                                        '¢Ð: Clear Error status

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
            
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "UPDATE  E11070T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " app_yn = " & FilterVar(UCase(arrColVal(7)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Updt_emp_no = " & FilterVar(Request("txtUpdtUserId"), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND convert(varchar(10),dilig_strt_dt,20) = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = " & FilterVar(UCase(arrColVal(6)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText
    if UCase(arrColVal(7)) = "Y" then
    '   ÀÜ¾÷table insert
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

        lgObjConn.Execute lgStrSQL,,adCmdText
    	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
    end if

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
                 .DBQueryOk()        
	          End with
	      Else
              With Parent
                 .DBQueryFail
	          End with
          End If   
       Case "UID_M0002"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
    End Select    
       
</Script>	
