<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    call LoadBasisGlobalInf()
                                                                  '☜: Clear Error status
    lgSvrDateTime = GetSvrDateTime
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
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
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If UNIConvDate(lgKeyStream(0)) = "" then
    Else
       strWhere = FilterVar(UNIConvDateCompanyToDB((lgKeyStream(0)),NULL),"NULL","S")
       strWhere = strWhere & " AND a.dilig_end_dt >= " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(0)),NULL),"NULL","S")
    End if

    If  lgKeyStream(1) = "" then
        strWhere = strWhere & " AND b.internal_cd  LIKE  " & FilterVar(Trim(lgKeyStream(2)) & "%", "''", "S") & " " 
    else
        strWhere = strWhere & " AND b.internal_cd  = " & FilterVar(lgKeyStream(2), "''", "S")
    end if
    
    If lgKeyStream(5) = "" then
       strWhere = strWhere & " AND a.dilig_cd LIKE " & "" & FilterVar("%", "''", "S") & ""
    Else
       strWhere = strWhere & " AND a.dilig_cd LIKE " & FilterVar(lgKeyStream(5), "''", "S")
    End if 
    
    If lgKeyStream(4) = "" then
       strWhere = strWhere & " AND a.emp_no LIKE " & "" & FilterVar("%", "''", "S") & ""
    Else
       strWhere = strWhere & " AND a.emp_no LIKE " & FilterVar(lgKeyStream(4), "''", "S")
    End if

    Call SubMakeSQLStatements("MR",strWhere,"X","<=")                                 '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ROLL_PSTN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ROLL_PSTN_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_NM"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("DILIG_STRT_DT"),"")
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("DILIG_END_DT"),"")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))
            
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
               
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

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
                    Call SubCheckHoliday(arrColVal)
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubCheckHoliday(arrColVal)
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
' Name : SubCheckHoliday
' Desc : Check Holiday
'============================================================================================================
Sub SubCheckHoliday(arrColVal)
	Dim strFg, strType, strOrgId
	Dim strHoli_type, strHoliday_apply
    Dim strWhere, strDilig_dt, strEnd_Dilig_dt
    Dim iCnt, iHoliday_cnt
    Dim IntRetCD
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

' 입력된일자의 근무당시의 부서코드검색 / 2002.04.08 송봉규 
    Call CommonQueryRs(" dept_cd "," hba010t a", " a.gazet_dt = (select MAX(gazet_dt) from hba010t where gazet_dt <= " & _
                                                                      FilterVar(UNIConvDate(arrColVal(2)), "''", "S") & _
	                                                              " and emp_no = a.emp_no" & ")" &_
		" and a.isrt_dt = (select MAX(isrt_dt) from hba010t where gazet_dt = a.gazet_dt " & _
	                                                              " and emp_no = a.emp_no" & ")" &_  	                                                              
                                                 " and a.emp_no = " & FilterVar(arrColVal(3), "''", "S") &_
                                                 " and a.dept_cd is not null",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strFg = Trim(Replace(lgF0, Chr(11), ""))

    If strFg = "" OR strFg = "X" Then
       Call CommonQueryRs(" dept_cd "," HAA010T "," emp_no = " & FilterVar(UCase(arrColVal(3)), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	   strFg = Replace(lgF0, Chr(11), "")
    End If
	
	strWhere = " a.dept_cd =  " & FilterVar(strFg , "''", "S") & "" & _
	           " AND a.org_change_dt = (SELECT MAX(org_change_dt) " &_
	                                   "  FROM b_acct_dept " &_
	                                   " WHERE dept_cd = a.dept_cd " &_
	                                   "   AND org_change_dt <= " & FilterVar(UNIConvDate(arrColVal(2)), "''", "S") & ")" &_
		       " AND a.cost_cd = b.cost_cd "                               
	Call CommonQueryRs(" b.biz_area_cd "," b_acct_dept a, b_cost_center b ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strOrgId = Replace(lgF0, Chr(11), "")

	If IsNull(strOrgId) or strOrgId = "" or strOrgId = "X" Then
        Call DisplayMsgBox("800486", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)        
        ObjectContext.SetAbort
        Call SetErrorStatus
		Exit Sub
	End If

	strWhere = " a.emp_no = " & FilterVar(UCase(arrColVal(3)), "''", "S") &_
	           " AND a.chang_dt = (SELECT MAX(chang_dt) " &_
	                            "  FROM hca040t " &_
	                            " WHERE emp_no = a.emp_no" &_
	                            "   AND chang_dt <= " & FilterVar(UNIConvDate(arrColVal(2)), "''", "S") & ")" 
  
    Call CommonQueryRs(" a.wk_type "," HCA040T a ", strWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	strType = Replace(lgF0, Chr(11), "")
	If (strType = "X" OR strType = "") Then strType = "0"

    strDilig_dt = UNIConvDate(arrColVal(5))
    strEnd_Dilig_dt = UNIConvDate(arrColVal(6))
    iCnt = 0
    iHoliday_cnt = 0
    
    '휴일적용여부가 'N'이고 해당일이 휴일이면 등록 불가 2002.11.06 by sbk 
    Call CommonQueryRs(" holiday_apply "," HCA010T "," dilig_cd = " & FilterVar(UCase(arrColVal(4)), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strHoliday_apply = Trim(Replace(lgF0, Chr(11), ""))
    
    '휴일적용여부가 'N'이고 기간내에 근태일이 모두 휴일이면 등록이 안되도록 함. 2002.11.08 by sbk
    If strHoliday_apply = "N" Then
        Do While strDilig_dt <= strEnd_Dilig_dt
            iCnt = iCnt +1

            '해당일이 휴일인지 평일인지 가져옴 2002.11.06 by sbk 
	        IntRetCD = CommonQueryRs(" holi_type "," HCA020T "," org_cd =  " & FilterVar(strOrgId , "''", "S") & "" & _
	                                 " and wk_type =  " & FilterVar(strType , "''", "S") & " and date =  " & FilterVar(strDilig_dt , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If IntRetCD = True Then
	            strHoli_type = Trim(Replace(lgF0, Chr(11), ""))
                If strHoli_type = "" OR strHoli_type = "X" Then
                    Call DisplayMsgBox("800453", vbInformation, "", "", I_MKSCRIPT)   
                    ObjectContext.SetAbort
                    Call SetErrorStatus
            		Exit Sub
                End If    
            Else
	        	Exit Sub
            End If

	        If strHoli_type = "H" Then 
                iHoliday_cnt = iHoliday_cnt +1
            Else
                Exit Do
	        End If

            strDilig_dt = UNIDateAdd("D", 1, strDilig_dt, gAPDateFormat)
        Loop 

        If iCnt = iHoliday_cnt Then
            Call DisplayMsgBox("800505", vbInformation, Trim(arrColVal(8)), "", I_MKSCRIPT)           
            ObjectContext.SetAbort
            Call SetErrorStatus
         	Exit Sub
	    End If
    End If
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
        
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HCA050T("
    lgStrSQL = lgStrSQL & " EMP_NO,"
    lgStrSQL = lgStrSQL & " DILIG_CD,"
    lgStrSQL = lgStrSQL & " DILIG_STRT_DT," 
    lgStrSQL = lgStrSQL & " DILIG_END_DT,"
    lgStrSQL = lgStrSQL & " REMARK," 
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(5)),"NULL","S")   & "," 
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(6)),"NULL","S")   & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(7), "''", "S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HCA050T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " DILIG_END_DT   = " & FilterVar(UNIConvDateCompanyToDB((arrColVal(6)),NULL),"NULL","S")   & "," 
    lgStrSQL = lgStrSQL & " REMARK  = "        & FilterVar(arrColVal(7), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE EMP_NO   = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND DILIG_CD = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND DILIG_STRT_DT = " & FilterVar(UNIConvDateCompanyToDB((arrColVal(5)),NULL),"NULL","S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HCA050T "
    lgStrSQL = lgStrSQL & " WHERE EMP_NO   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND DILIG_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND DILIG_STRT_DT = " & FilterVar(UNIConvDateCompanyToDB((arrColVal(4)),NULL),"NULL","S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount  & " b.name, a.EMP_NO, a.DILIG_STRT_DT, a.DILIG_END_DT, a.REMARK, "
                       lgStrSQL = lgStrSQL & " b.roll_pstn, dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",b.ROLL_PSTN) ROLL_PSTN_NM, "
                       lgStrSQL = lgStrSQL & " a.DILIG_CD, dbo.ufn_H_GetCodeName(" & FilterVar("HCA010T", "''", "S") & ",a.DILIG_CD,'') DILIG_NM "
                       lgStrSQL = lgStrSQL & " From  HCA050T a, HAA010T b, HCA010T c "
                       lgStrSQL = lgStrSQL & " Where a.emp_no = b.emp_no AND c.dilig_cd = a.dilig_cd  AND c.day_time=" & FilterVar("1", "''", "S") & "  AND a.dilig_strt_dt " & pComp & pCode
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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
