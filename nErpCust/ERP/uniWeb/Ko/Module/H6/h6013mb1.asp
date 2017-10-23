<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim strWhere
    Dim strSub_type   
    Dim strPAY_CD  
    Dim strDept_nm
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strWhere = FilterVar(lgKeyStream(0), "''", "S") 
    strWhere = strWhere & " AND c.emp_no LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ""   
    strWhere = strWhere & " AND b.pay_cd LIKE  " & FilterVar(lgKeyStream(2) & "%", "''", "S") & ""
    strWhere = strWhere & " AND c.prov_type LIKE  " & FilterVar(lgKeyStream(3) & "%", "''", "S") & ""
    strWhere = strWhere & " AND c.allow_cd LIKE  " & FilterVar(lgKeyStream(4) & "%", "''", "S") & ""
    strWhere = strWhere & " AND b.internal_cd LIKE   " & FilterVar(lgKeyStream(6) & "%", "''", "S") & ""
'    strWhere = strWhere & " and c.emp_no = a.emp_no "
    strWhere = strWhere & " and c.emp_no = b.emp_no "
    strWhere = strWhere & " and c.emp_no = d.emp_no "
    strWhere = strWhere & " and c.pay_yymm = d.pay_yymm "
    strWhere = strWhere & " and c.prov_type = d.prov_type "
    
    If lgKeyStream(5) <> 0 Then
        strWhere = strWhere & " AND c.allow >= " & UNIConvNum(lgKeyStream(5), 0)
    End If
            
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
         
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("NAME")))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("EMP_NO")))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("DEPT_NM")))
           
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("ALLOW_CD_NM")))   '수당코드 
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ALLOW"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE_NM"))
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
                       lgStrSQL = "Select TOP " & iSelCount
                       lgStrSQL = lgStrSQL & " b.name     ,"   
                       lgStrSQL = lgStrSQL & " c.EMP_NO   ,"   
                       lgStrSQL = lgStrSQL & " d.DEPT_CD  , d.DEPT_NM ," 
                       lgStrSQL = lgStrSQL & " dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ", c.ALLOW_CD, '') ALLOW_CD_NM,"   
                       lgStrSQL = lgStrSQL & " c.ALLOW    ,"   
                       lgStrSQL = lgStrSQL & " b.PAY_CD   , dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ", b.PAY_CD) PAY_CD_NM,"   
                       lgStrSQL = lgStrSQL & " c.PROV_TYPE, dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", c.PROV_TYPE) PROV_TYPE_NM " 
                       lgStrSQL = lgStrSQL & " From  HDF040T c ,HDF020T b ,HDF070T d"
                       lgStrSQL = lgStrSQL & " Where c.PAY_YYMM " & pComp & pCode
                       lgStrSQL = lgStrSQL & " Order By d.INTERNAL_CD, c.EMP_NO "
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
