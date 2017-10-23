<% Option Explicit%>
<%Response.Buffer = True%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    
    Dim lgSvrDateTime
    lgSvrDateTime = GetSvrDateTime
        
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
     End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim strWhere
    Dim strWhereHead
    Dim strWhereTab1
    Dim strWhereTab2
    Dim strWherefooter
    Dim txtFrom_dt
    Dim txtTo_dt
    Dim txtFr_internal_cd
    Dim txtTo_internal_cd
    Dim txtPay_grd1
    Dim gSelframeFlg
    Dim stDT, endT, baseDt
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
 
	if lgKeyStream(6) = "" then
		stDT = "1900-01-01"
	else
		stDT = uniGetFirstDay(lgKeyStream(6),gDateFormatYYYYMM)
		stDT = UniConvDate(stDt)
	end if

	if lgKeyStream(7) = "" then
		endT = "2500-12-01"
	else
		enDT = uniGetFirstDay(lgKeyStream(7),gDateFormatYYYYMM)
		enDT = UniConvDate(enDT)
	end if

	baseDt = FilterVar(lgKeyStream(9),"'%'", "S")

    strWhere = " ( dbo.ufn_H_get_internal_cd (emp_no,"& baseDt &" ) >=  " & FilterVar(lgKeyStream(1), "''", "S")
    strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd (emp_no,"& baseDt &" ) <= " & FilterVar(lgKeyStream(2), "''", "S") & ")"

    If lgKeyStream(0) <> "" then
       strWhere =  strWhere & " and emp_no = " &  FilterVar(lgKeyStream(0), "''", "S") 
    End if
    
	strWhere = strWhere & " AND ((APPLY_YYMM IS NULL AND REVOKE_YYMM IS  NULL)"
	strWhere = strWhere & "   OR (APPLY_YYMM IS NULL AND REVOKE_YYMM IS NOT NULL AND REVOKE_YYMM >=   " & FilterVar(stDT , "''", "S") & ") "
	strWhere = strWhere & "   OR (REVOKE_YYMM IS NULL AND APPLY_YYMM IS NOT NULL AND APPLY_YYMM <=  " & FilterVar(enDT , "''", "S") & ") "
	strWhere = strWhere & "   OR ((APPLY_YYMM BETWEEN  " & FilterVar(stDT, "''", "S") & " and  " & FilterVar(enDT, "''", "S") & ") OR (REVOKE_YYMM BETWEEN   " & FilterVar(stDT, "''", "S") & " and  " & FilterVar(enDT, "''", "S") & ")) "
	strWhere = strWhere & "   OR (( " & FilterVar(stDT, "''", "S") & " BETWEEN APPLY_YYMM AND REVOKE_YYMM) OR ( " & FilterVar(enDT, "''", "S") & " BETWEEN APPLY_YYMM AND REVOKE_YYMM))) "

    If lgKeyStream(3) <> "" then
       strWhere = strWhere & " AND  ( allow_cd LIKE " & FilterVar(lgKeyStream(3), "''", "S") & ")"
    End if 

    If lgKeyStream(8) = "Y" then
		strWhere = strWhere & " AND ALLOW_TYPE = 'Y' "
    ElseIf lgKeyStream(8) = "N" then
		strWhere = strWhere & " AND (ALLOW_TYPE is null or ALLOW_TYPE = 'N') "
    End if
    
    Call SubMakeSQLStatements("MR",strWhere,baseDt,C_EQGT)                      '¡Ù : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_NM"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ALLOW"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNIMonthClientFormat(lgObjRs("APPLY_YYMM"))
            lgstrData = lgstrData & Chr(11) & UNIMonthClientFormat(lgObjRs("REVOKE_YYMM"))
            
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
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

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
'-----------
	Dim itxtSpread
	Dim itxtSpreadArr
	Dim itxtSpreadArrCount
	Dim iCUCount
	Dim iDCount
	Dim ii
	
	itxtSpread = ""
	             
	iCUCount = Request.Form("txtCUSpread").Count
	iDCount  = Request.Form("txtDSpread").Count

	itxtSpreadArrCount = -1
	             
	ReDim itxtSpreadArr(iCUCount + iDCount)
	             
	For ii = 1 To iDCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
	Next
	For ii = 1 To iCUCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
	Next
 
   itxtSpread = Join(itxtSpreadArr,"")
	arrRowVal = Split(itxtSpread, gRowSep)                                 '¢Ð: Split Row    data   
 '---------        
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
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
    Dim Apply_Yymm, Revoke_Yymm
    Dim strYear, strMonth, strDay

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	If Trim(arrColVal(7)) = gComDateType OR Trim(arrColVal(7)) = "" Then
		Apply_Yymm = ""
	Else
		Call ExtractDateFrom(arrColVal(7), gDateFormatYYYYMM, gComDateType, strYear, strMonth, strDay)
		Apply_Yymm = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")
	End If

	If Trim(arrColVal(8)) = gComDateType OR Trim(arrColVal(8)) = "" Then
		Revoke_Yymm = ""
	Else
		Call ExtractDateFrom(arrColVal(8), gDateFormatYYYYMM, gComDateType, strYear, strMonth, strDay)
		Revoke_Yymm = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")
	End If
    
    lgStrSQL = "INSERT INTO HDF030T("
    lgStrSQL = lgStrSQL & " EMP_NO," 
    lgStrSQL = lgStrSQL & " ALLOW_CD," 
    lgStrSQL = lgStrSQL & " ALLOW," 
    lgStrSQL = lgStrSQL & " ALLOW_TYPE,"
    lgStrSQL = lgStrSQL & " APPLY_YYMM," 
    lgStrSQL = lgStrSQL & " REVOKE_YYMM,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO," 
    lgStrSQL = lgStrSQL & " UPDT_DT)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0)     & ","
    lgStrSQL = lgStrSQL & " " & FilterVar("Y", "''", "S") & " ,"
    lgStrSQL = lgStrSQL & FilterVar(Apply_Yymm,"NULL","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Revoke_Yymm,"NULL","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"
'Response.Write lgStrSQL
'Response.End 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    Dim Apply_Yymm, Revoke_Yymm
    Dim strYear, strMonth, strDay

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	If Trim(arrColVal(5)) = gComDateType OR Trim(arrColVal(5)) = "" Then
		Apply_Yymm = ""
	Else
		Call ExtractDateFrom(arrColVal(5), gDateFormatYYYYMM, gComDateType, strYear, strMonth, strDay)
		Apply_Yymm = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")
	End If

	If Trim(arrColVal(6)) = gComDateType OR Trim(arrColVal(6)) = "" Then
		Revoke_Yymm = ""
	Else
		Call ExtractDateFrom(arrColVal(6), gDateFormatYYYYMM, gComDateType, strYear, strMonth, strDay)
		Revoke_Yymm = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")
	End If

    lgStrSQL = "UPDATE  HDF030T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " ALLOW        = " & UNIConvNum(arrColVal(4),0) & ", "
    lgStrSQL = lgStrSQL & " APPLY_YYMM   = " & FilterVar(Apply_Yymm,"NULL","S")    & ","
    lgStrSQL = lgStrSQL & " REVOKE_YYMM  = " & FilterVar(Revoke_Yymm,"NULL","S")
    lgStrSQL = lgStrSQL & " WHERE EMP_NO   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND ALLOW_CD  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
'Response.Write lgStrSQL
'Response.End    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "DELETE  HDF030T "
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " EMP_NO   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND ALLOW_CD  = " & FilterVar(UCase(arrColVal(3)), "''", "S")

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
						lgStrSQL = "SELECT TOP " & iSelCount  & " EMP_NO, ALLOW_CD, ALLOW, APPLY_YYMM, REVOKE_YYMM, ALLOW_TYPE, dbo.ufn_H_GetEmpName(emp_no) name, dbo.ufn_H_GetCodeName('HDA010T',ALLOW_CD,'') allow_nm "
						lgStrSQL = lgStrSQL & " ,dbo.ufn_H_get_dept_cd (EMP_NO ," & pCode1 & ") dept_cd"
						lgStrSQL = lgStrSQL & "	,dbo.ufn_GetDeptName(dbo.ufn_H_get_dept_cd (EMP_NO , " & pCode1 & " )," & pCode1 & ") dept_nm"                       
						lgStrSQL = lgStrSQL & "  FROM HDF030T  "
						lgStrSQL = lgStrSQL & " WHERE "  &  pCode
    					lgStrSQL = lgStrSQL & " ORDER BY dbo.ufn_H_get_dept_cd (EMP_NO ," & pCode1 & ") ASC, EMP_NO ASC "
'Response.Write lgStrSQL
'Response.End
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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
       
</Script>	
