<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

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
    Dim strWhere
    Dim strDept_nm
    Dim stDT,enDT
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	if lgKeyStream(5) = "" then
		stDT = "1900-01-01"
	else
		stDT = uniGetFirstDay(lgKeyStream(5),gDateFormatYYYYMM)
		stDT = UniConvDate(stDt)
	end if
	
	if lgKeyStream(6) = "" then
		endT = "2500-12-31"
	else
		enDT = uniGetLastDay(lgKeyStream(6),gDateFormatYYYYMM)
		enDT = UniConvDate(enDT)
	end if
 
    strWhere = " LIKE " & FilterVar(lgKeyStream(0),"'%'", "S")
    strWhere = strWhere & " AND B.SUB_TYPE LIKE " & FilterVar(lgKeyStream(1),"'%'", "S")
    strWhere = strWhere & " AND B.SUB_CD LIKE " & FilterVar(lgKeyStream(2),"'%'", "S")
    strWhere = strWhere & " AND A.INTERNAL_CD >=  " & FilterVar(lgKeyStream(3), "''", "S") & ""       '  internal_cd = min
    strWhere = strWhere & " AND A.INTERNAL_CD <=  " & FilterVar(lgKeyStream(4), "''", "S") & ""       '  internal_cd = max
    strWhere = strWhere & " AND A.INTERNAL_CD LIKE  " & FilterVar(lgKeyStream(9) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'
    strWhere = strWhere & " AND A.EMP_NO = B.EMP_NO "  

	strWhere = strWhere & " AND ((APPLY_YYMM IS NULL AND REVOKE_YYMM IS  NULL)"
	strWhere = strWhere & "   OR (APPLY_YYMM IS NULL AND REVOKE_YYMM IS NOT NULL AND REVOKE_YYMM >=   " & FilterVar(stDT , "''", "S") & ") "
	strWhere = strWhere & "   OR (REVOKE_YYMM IS NULL AND APPLY_YYMM IS NOT NULL AND APPLY_YYMM <=  " & FilterVar(enDT , "''", "S") & ") "
	strWhere = strWhere & "   OR ((APPLY_YYMM BETWEEN  " & FilterVar(stDT, "''", "S") & " and  " & FilterVar(enDT, "''", "S") & ") OR (REVOKE_YYMM BETWEEN   " & FilterVar(stDT, "''", "S") & " and  " & FilterVar(enDT, "''", "S") & ")) "
	strWhere = strWhere & "   OR (( " & FilterVar(stDT, "''", "S") & " BETWEEN APPLY_YYMM AND REVOKE_YYMM) OR ( " & FilterVar(enDT, "''", "S") & " BETWEEN APPLY_YYMM AND REVOKE_YYMM)))  "

    strWhere = strWhere & " ORDER BY A.DEPT_CD , B.EMP_NO "

    Call SubMakeSQLStatements("MR",strWhere,"X","")                                 '☆ : Make sql statements

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
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_CD_NM"))
            lgstrData = lgstrData & Chr(11) & ""
          	lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("SUB_AMT"), ggAmtOfMoney.DecPoint,0)
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
	arrRowVal = Split(itxtSpread, gRowSep)                                 '☜: Split Row    data   
 '---------        

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
    Dim Apply_Yymm, Revoke_Yymm
    Dim strYear, strMonth, strDay
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call ExtractDateFrom(arrColVal(8), gDateFormatYYYYMM, gComDateType, strYear, strMonth, strDay)
    Apply_Yymm = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")
    Call ExtractDateFrom(arrColVal(9), gDateFormatYYYYMM, gComDateType, strYear, strMonth, strDay)
    Revoke_Yymm = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")
    
    lgStrSQL = "INSERT INTO HDF050T("
    lgStrSQL = lgStrSQL & " EMP_NO," 
    lgStrSQL = lgStrSQL & " SUB_TYPE," 
    lgStrSQL = lgStrSQL & " SUB_CD,"
    lgStrSQL = lgStrSQL & " SUB_AMT,"
    lgStrSQL = lgStrSQL & " APPLY_YYMM," 
    lgStrSQL = lgStrSQL & " REVOKE_YYMM," 
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO ," 
    lgStrSQL = lgStrSQL & " ISRT_DT ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO ," 
    lgStrSQL = lgStrSQL & " UPDT_DT )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(7)),0)     & ","
    lgStrSQL = lgStrSQL & FilterVar(Apply_Yymm,"NULL","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Revoke_Yymm,"NULL","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"

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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call ExtractDateFrom(arrColVal(8), gDateFormatYYYYMM, gComDateType, strYear, strMonth, strDay)
    Apply_Yymm = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")
    Call ExtractDateFrom(arrColVal(9), gDateFormatYYYYMM, gComDateType, strYear, strMonth, strDay)
    Revoke_Yymm = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")
        
    lgStrSQL = "UPDATE  HDF050T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " SUB_AMT      = " & UNIConvNum(Trim(arrColVal(7)),0)     & ","
    lgStrSQL = lgStrSQL & " APPLY_YYMM   = " & FilterVar(Apply_Yymm,"NULL","S")    & ","
    lgStrSQL = lgStrSQL & " REVOKE_YYMM  = " & FilterVar(Revoke_Yymm,"NULL","S")
    lgStrSQL = lgStrSQL & " WHERE    "
    lgStrSQL = lgStrSQL & " EMP_NO       = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " And SUB_TYPE = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    lgStrSQL = lgStrSQL & " And SUB_CD   = " & FilterVar(UCase(arrColVal(6)), "''", "S")

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

    lgStrSQL = "DELETE  HDF050T"
    lgStrSQL = lgStrSQL & " WHERE    "
    lgStrSQL = lgStrSQL & " EMP_NO       = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND SUB_TYPE     = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    lgStrSQL = lgStrSQL & " AND SUB_CD       = " & FilterVar(UCase(arrColVal(6)), "''", "S")
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
                       lgStrSQL = "SELECT TOP " & iSelCount
                       lgStrSQL = lgStrSQL & " A.NAME , B.EMP_NO ,"   
                       lgStrSQL = lgStrSQL & " A.DEPT_CD     ,"   
                       lgStrSQL = lgStrSQL & " A.DEPT_NM     ,"   
                       lgStrSQL = lgStrSQL & " B.SUB_TYPE    ,"   
                       lgStrSQL = lgStrSQL & " dbo.ufn_getCodeName(" & FilterVar("H0040", "''", "S") & ", B.SUB_TYPE) SUB_TYPE_NM ,"   
                       lgStrSQL = lgStrSQL & " B.SUB_CD      ,"   
                       lgStrSQL = lgStrSQL & " dbo.ufn_H_GetCodeName(" & FilterVar("HDA010t", "''", "S") & ", B.SUB_CD ,'') SUB_CD_NM, "
                       lgStrSQL = lgStrSQL & " b.SUB_AMT     ,"   
                       lgStrSQL = lgStrSQL & " B.APPLY_YYMM  ,"   
                       lgStrSQL = lgStrSQL & " B.REVOKE_YYMM ," 
                       lgStrSQL = lgStrSQL & " B.CALCU_TYPE  ,"   
                       lgStrSQL = lgStrSQL & " B.ISRT_EMP_NO ,"   
                       lgStrSQL = lgStrSQL & " B.ISRT_DT     ,"   
                       lgStrSQL = lgStrSQL & " B.UPDT_EMP_NO ,"   
                       lgStrSQL = lgStrSQL & " B.UPDT_DT, "
                       lgStrSQL = lgStrSQL & " A.INTERNAL_CD "
                       lgStrSQL = lgStrSQL & " FROM HAA010T A, HDF050T B "
                       lgStrSQL = lgStrSQL & " WHERE B.EMP_NO " & pComp & pCode
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

