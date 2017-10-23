<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    
    Call LoadBasisGlobalInf()
    Call LoadinfTb19029B("Q", "H","NOCOOKIE","MB")                                                                   '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    
    lgPrevNext        = Request("txtPrevNext") 

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)  
            Call SubBizQuery()                                                          '☜: Query
            Call SubBizQueryMulti()
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
    Dim strWhere,rDate
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strWhere = FilterVar(lgKeyStream(0), "''", "S")
    strWhere = strWhere & " AND HGA080T.EMP_NO LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ""
    strWhere = strWhere & " AND HAA010T.emp_no = HGA080T.EMP_NO "
    strWhere = strWhere & " AND HGA080T.DEPT_CD = B_ACCT_DEPT.DEPT_CD "
    strWhere = strWhere & " AND B_ACCT_DEPT.ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) "
    strWhere = strWhere & "                                  FROM B_ACCT_DEPT "
    strWhere = strWhere & "                                  WHERE ORG_CHANGE_DT <=  " & FilterVar(UNIConvDateCompanyToDB(lgKeyStream(3),NULL), "''", "S") & ")"  ' 월의 마지막 날자    
    strWhere = strWhere & " AND HGA080T.internal_cd LIKE   " & FilterVar(lgKeyStream(2) & "%", "''", "S") & ""
    strWhere = strWhere & " Order by HGA080T.EMP_NO "
    
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                              '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '☜ : No data is found.
%>        
        <Script Language="VBScript">
        Parent.Nodata
        </script>
<%
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TOT_DUTY_MM"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_ESTI_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BONUS_ESTI_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("YEAR_ESTI_AMT"), ggAmtOfMoney.DecPoint,0)
            
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("AVR_WAGES"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RETIRE_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TOT_PROV_AMT"), ggAmtOfMoney.DecPoint,0)
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DUTY_YY"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DUTY_MM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DUTY_DD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RETIRE_YYMM"))

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
                       lgStrSQL = "Select TOP " & iSelCount 
                       lgStrSQL = lgStrSQL & " HGA080T.EMP_NO, HGA080T.DEPT_CD, HGA080T.TOT_DUTY_MM, B_ACCT_DEPT.DEPT_NM,"   
                       lgStrSQL = lgStrSQL & " HGA080T.PAY_ESTI_AMT, HGA080T.BONUS_ESTI_AMT, HGA080T.TOT_PROV_AMT,"   
                       lgStrSQL = lgStrSQL & " HAA010T.name, HGA080T.RETIRE_YYMM, HGA080T.DUTY_YY, HGA080T.DUTY_MM,"   
                       lgStrSQL = lgStrSQL & " HGA080T.DUTY_DD, HGA080T.AVR_WAGES, HGA080T.RETIRE_AMT, HGA080T.HONOR_AMT, "   
                       lgStrSQL = lgStrSQL & " HGA080T.ETC_AMT, HGA080T.YEAR_ESTI_AMT, HGA080T.UPDT_EMP_NO, HGA080T.UPDT_DT "
                       lgStrSQL = lgStrSQL & " FROM HGA080T,HAA010T,B_ACCT_DEPT "  
                       lgStrSQL = lgStrSQL & " WHERE HGA080T.RETIRE_YYMM " & pComp & pCode
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
