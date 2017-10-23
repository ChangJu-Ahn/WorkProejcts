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
    'Call svrmsgbox(lgKeyStream(7),0,1)
    strWhere = FilterVar(lgKeyStream(0), "''", "S")  'frm1.pau_yymm
    strWhere = strWhere & " AND a.emp_no LIKE " & FilterVar(Trim(lgKeyStream(1)),"" & FilterVar("%", "''", "S") & "","S")   
    strWhere = strWhere & " AND b.pay_cd LIKE " & FilterVar(Trim(lgKeyStream(2)),"" & FilterVar("%", "''", "S") & "","S")    
    strWhere = strWhere & " AND b.prov_type LIKE " & FilterVar(Trim(lgKeyStream(3)),"" & FilterVar("%", "''", "S") & "","S")   
    
'    strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(7),"'%'", "S") & ") >= " & FilterVar(lgKeyStream(4), "''", "S")
'    strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(7),"'%'", "S") & ") <= " & FilterVar(lgKeyStream(5), "''", "S")
'    strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(7),"'%'", "S") & ") LIKE " & FilterVar(lgKeyStream(6) & "%", "''", "S") 
    
    strWhere = strWhere & " AND b.internal_cd >=  " & FilterVar(lgKeyStream(4), "''", "S")					'  internal_cd = min
    strWhere = strWhere & " AND b.internal_cd <=  " & FilterVar(lgKeyStream(5), "''", "S")					'  internal_cd = max
    strWhere = strWhere & " AND b.internal_cd LIKE  " & FilterVar(lgKeyStream(6) & "%", "''", "S")			'  자료권한 Internal_cd = '?'
    strWhere = strWhere & " and a.emp_no = b.emp_no "
           
    Call SubMakeSQLStatements("MR",strWhere,"X","=")                                 '☆ : Make sql statements

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
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs("EMP_NO")) )           
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD_NM"))  '급여구분 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE_NM"))  '지급구분 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("INCOME_TAX"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RES_TAX"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ANUT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MED_INSUR"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("EMP_INSUR"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("SUB_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("REAL_PROV_AMT"), ggAmtOfMoney.DecPoint,0)
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
                       lgStrSQL = lgStrSQL & " a.NAME         ,"   
                       lgStrSQL = lgStrSQL & " a.EMP_NO       ,"   
                       lgStrSQL = lgStrSQL & " b.DEPT_CD      ,"   
                       lgStrSQL = lgStrSQL & " b.DEPT_NM      ,"   
                       lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ", b.PAY_CD) PAY_CD_NM,"   
                       lgStrSQL = lgStrSQL & " CASE b.PROV_TYPE WHEN " & FilterVar("1", "''", "S") & "  THEN b.PAY_TOT_AMT WHEN " & FilterVar("P", "''", "S") & "  THEN b.PAY_TOT_AMT ELSE b.BONUS_TOT_AMT END AS PAY_TOT_AMT ,"   
                       lgStrSQL = lgStrSQL & " b.INCOME_TAX   ,"   
                       lgStrSQL = lgStrSQL & " b.RES_TAX      ,"   
                       lgStrSQL = lgStrSQL & " b.ANUT         ,"   
                       lgStrSQL = lgStrSQL & " b.MED_INSUR    ,"   
                       lgStrSQL = lgStrSQL & " b.EMP_INSUR    ,"   
                       lgStrSQL = lgStrSQL & " b.SUB_TOT_AMT  ,"   
                       lgStrSQL = lgStrSQL & " b.REAL_PROV_AMT,"   
                       lgStrSQL = lgStrSQL & " b.PROV_TYPE, dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", b.PROV_TYPE) PROV_TYPE_NM "  
                       lgStrSQL = lgStrSQL & " From HDF020T a ,HDF070T b"
                       lgStrSQL = lgStrSQL & " Where b.PAY_YYMM " & pComp & pCode
                       lgStrSQL = lgStrSQL & " order by b.dept_cd, a.emp_no,b.PROV_TYPE "
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
