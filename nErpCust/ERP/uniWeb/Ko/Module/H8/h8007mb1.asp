<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim strWhere
    Dim strpay_grd1_nm
    Dim strpay_grd1_nm1
    Dim strdept_cd
    Dim strdept_nm
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strWhere = FilterVar(lgKeyStream(0), "''", "S") 
    
    strWhere = strWhere & " AND HDF070T.internal_cd >=  " & FilterVar(lgKeyStream(1), "''", "S") & ""       '  internal_cd = min
    strWhere = strWhere & " AND HDF070T.internal_cd <=  " & FilterVar(lgKeyStream(2), "''", "S") & ""       '  internal_cd = max
    strWhere = strWhere & " AND HDF070T.internal_cd LIKE  " & FilterVar(lgKeyStream(3) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'
    
    strWhere = strWhere & " and HDF040T.prov_type IN (" & FilterVar("1", "''", "S") & " ," & FilterVar("P", "''", "S") & " ," & FilterVar("B", "''", "S") & " ) "
    strWhere = strWhere & " and HDF040T.emp_no = HDF070T.emp_no "
    strWhere = strWhere & " and HDF040T.pay_yymm = HDF070T.pay_yymm "
    strWhere = strWhere & " and HDF040T.prov_type = HDF070T.prov_type "
    strWhere = strWhere & " and HDA010T.code_type = " & FilterVar("1", "''", "S") & "  "
    strWhere = strWhere & " and HDA010T.allow_cd = HDF040T.allow_cd "
    strWhere = strWhere & " and HDF070T.emp_no IN (SELECT emp_no FROM hdf070t WHERE pay_yymm = " & FilterVar(lgKeyStream(0), "''", "S")
    strWhere = strWhere & " and prov_type = " & FilterVar("P", "''", "S") & " ) "
    strWhere = strWhere & " GROUP BY dept_nm , pay_grd1 , HDF040T.allow_cd "
    strWhere = strWhere & " with rollup "

	Call SubMakeSQLStatements("MR",strWhere,"X","=")                                 '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        iDx       = 1
        strpay_grd1_nm1 = ""
            
        Do While Not lgObjRs.EOF
        
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & lgObjRs("DEPT_CD")
                             
            If lgObjRs("DEPT_CD") = "총계" Then
                lgstrData = lgstrData & Chr(11) & ""
            Else
                lgstrData = lgstrData & Chr(11) & lgObjRs("PAY_GRD1")
            End If
            
            If lgObjRs("PAY_GRD1") = "합계" Then
                lgstrData = lgstrData & Chr(11) & ""
            Else
                lgstrData = lgstrData & Chr(11) & lgObjRs("ALLOW_CD")
            End If
            
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ORIGINAL_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RAISE_AMT")   , ggAmtOfMoney.DecPoint,0) 
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RETRO_AMT")   , ggAmtOfMoney.DecPoint,0) 
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
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
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
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount
  
    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
           
                       lgStrSQL = "Select CASE WHEN (GROUPING(HDF070T.DEPT_NM) = 1) THEN " & FilterVar("총계", "''", "S") & "  ELSE HDF070T.DEPT_NM END AS dept_cd , "
                       lgStrSQL = lgStrSQL & " CASE WHEN (GROUPING(HDF070T.PAY_GRD1) = 1) THEN " & FilterVar("합계", "''", "S") & "  ELSE dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",HDF070T.PAY_GRD1) END AS pay_grd1 , "
                       lgStrSQL = lgStrSQL & " CASE WHEN (GROUPING(HDF040T.ALLOW_CD) = 1) THEN " & FilterVar("소계", "''", "S") & "  ELSE dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ",HDF040T.ALLOW_CD,'') END AS allow_cd , "
                       lgStrSQL = lgStrSQL & " convert(numeric(18,4), sum(case HDF040T.prov_type when " & FilterVar("1", "''", "S") & "  then isnull(HDF040T.ALLOW,0) else 0 end)) original_amt, "
                       lgStrSQL = lgStrSQL & " convert(numeric(18,4), sum(case HDF040T.prov_type when " & FilterVar("P", "''", "S") & "  then isnull(HDF040T.ALLOW,0) else 0 end)) raise_amt, "
                       lgStrSQL = lgStrSQL & " convert(numeric(18,4), sum(case HDF040T.prov_type when " & FilterVar("P", "''", "S") & "  then isnull(HDF040T.ALLOW,0) else 0 end)) "
                       lgStrSQL = lgStrSQL & " - convert(numeric(18,4), sum(case HDF040T.prov_type when " & FilterVar("1", "''", "S") & "  then isnull(HDF040T.ALLOW,0) else 0 end)) retro_amt "
                       lgStrSQL = lgStrSQL & " From HDF040T ,HDF070T ,HDA010T "
                       lgStrSQL = lgStrSQL & " Where HDF040T.PAY_YYMM " & pComp & pCode
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
