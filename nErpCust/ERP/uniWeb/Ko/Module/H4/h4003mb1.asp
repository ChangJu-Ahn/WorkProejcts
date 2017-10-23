<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
	Dim	lgStrPrevKey, lgSvrDateTime
	Const C_SHEETMAXROWS_D = 100

    call LoadBasisGlobalInf()
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")
        
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

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
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim strWhere
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    If lgKeyStream(0) = "" then
    Else
       strWhere = FilterVar(UNIConvDateCompanyToDB((lgKeyStream(0)),NULL),"NULL","S")
    End if
    
    If  lgKeyStream(1) = "" then
        strWhere = strWhere & " AND a.INTERNAL_CD  LIKE  " & FilterVar(Trim(lgKeyStream(8)) & "%", "''", "S") & " " 
    else
        strWhere = strWhere & " AND a.INTERNAL_CD  = " & FilterVar(lgKeyStream(8), "''", "S")
    end if
    
    If lgKeyStream(3) = "" then
        strWhere = strWhere & " AND b.emp_no LIKE " & "" & FilterVar("%", "''", "S") & ""
    Else
       strWhere = strWhere & " AND b.emp_no LIKE " & FilterVar(lgKeyStream(3), "''", "S")
    End if 

    Call SubMakeSQLStatements("MR",strWhere,"X","=")                                 '¡Ù : Make sql statements
    
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
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD_NM"))
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("STRT_DATE"),"")
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("STRT_HH"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("STRT_MM"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNIConvDateDBToCompany(lgObjRs("END_DATE"),"")
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("END_HH"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("END_MM"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("HOLI_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("HOLI_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INTERNAL_CD"))
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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "INSERT INTO HCA080T("
    lgStrSQL = lgStrSQL & " ATTEND_DT," 
    lgStrSQL = lgStrSQL & " EMP_NO," 
    lgStrSQL = lgStrSQL & " DEPT_CD," 
    lgStrSQL = lgStrSQL & " WK_TYPE," 
    lgStrSQL = lgStrSQL & " HOLI_TYPE,"
    lgStrSQL = lgStrSQL & " STRT_DATE,"
    lgStrSQL = lgStrSQL & " STRT_HH," 
    lgStrSQL = lgStrSQL & " STRT_MM," 
    lgStrSQL = lgStrSQL & " END_DATE," 
    lgStrSQL = lgStrSQL & " END_HH,"
    lgStrSQL = lgStrSQL & " END_MM," 
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO ," 
    lgStrSQL = lgStrSQL & " UPDT_DT ," 
    lgStrSQL = lgStrSQL & " INTERNAL_CD)" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB((arrColVal(2)),NULL),"NULL","S")   & "," 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB((arrColVal(7)),NULL),"NULL","S")   & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(8),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0)     & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB((arrColVal(10)),NULL),"NULL","S")   & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(11),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(12),0)     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,NULL,"S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(13)), "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
   
    lgStrSQL = "UPDATE  HCA080T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " STRT_DATE = "	& FilterVar(UNIConvDateCompanyToDB((arrColVal(4)),NULL),"NULL","S")	& "," 
    lgStrSQL = lgStrSQL & " STRT_HH = "		&  UNIConvNum(arrColVal(5),0)                                           & ","
    lgStrSQL = lgStrSQL & " STRT_MM	= "		&  UNIConvNum(arrColVal(6),0)                                           & ","
    lgStrSQL = lgStrSQL & " END_DATE= "		& FilterVar(UNIConvDateCompanyToDB((arrColVal(7)),NULL),"NULL","S")	& "," 
    lgStrSQL = lgStrSQL & " END_HH = "		&  UNIConvNum(arrColVal(8),0)                                           & ","
    lgStrSQL = lgStrSQL & " END_MM = "		&  UNIConvNum(arrColVal(9),0)                                           & ","
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO = "	& FilterVar(gUsrId, "''", "S")												& "," 
    lgStrSQL = lgStrSQL & " UPDT_DT = "		& FilterVar(lgSvrDateTime,NULL,"S")
    
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " ATTEND_DT   = " & FilterVar(UNIConvDateCompanyToDB((arrColVal(2)),NULL),"NULL","S")
    lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    
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

    lgStrSQL = "DELETE  HCA080T "
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " ATTEND_DT   = " & FilterVar(UNIConvDateCompanyToDB((arrColVal(2)),NULL),"NULL","S")
    lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    
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
                       lgStrSQL = "Select TOP " & iSelCount  & " a.EMP_NO,NAME,a.DEPT_CD,STRT_DATE,"
                       lgStrSQL = lgStrSQL & "STRT_HH,STRT_MM,END_DATE,END_HH,END_MM,WK_TYPE,"
                       lgStrSQL = lgStrSQL & "HOLI_TYPE, a.INTERNAL_CD,"
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetDeptName(b.DEPT_CD," &FilterVar(UNIConvDateCompanyToDB((lgKeyStream(0)),NULL),"NULL","S")  & ") DEPT_CD_NM,"
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0047", "''", "S") & ",WK_TYPE) WK_TYPE_NM,"
                       lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName(" & FilterVar("H0049", "''", "S") & ",HOLI_TYPE) HOLI_TYPE_NM"
                       lgStrSQL = lgStrSQL & " From  HAA010T a, HCA080T b "
                       lgStrSQL = lgStrSQL & " Where a.emp_no = b.emp_no AND b.attend_dt " & pComp & pCode
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
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
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
