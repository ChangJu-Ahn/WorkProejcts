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
Sub SubBizQuery()
    Dim iDx
    Dim iLoopMax
    Dim strWhere
    Dim txtFrom_dt , txtFrom_dt2
    Dim txtTo_dt , txtTo_dt2
    Dim txtsub_cd2
    Dim txtAllow_cd
    Dim txtOcpt_type , txtOcpt_type2
    Dim txtSect_cd , txtSect_cd2
    Dim txtFr_internal_cd , txtFr_internal_cd2
    Dim txtTo_internal_cd , txtTo_internal_cd2
    Dim rbo_sort
    Dim txtsub_type2
    Dim txtPayCd
    Dim FrSQL,ToSQL
    Dim FrSQL2,ToSQL2 
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    txtFrom_dt          = FilterVar(lgKeyStream(1), "''", "S")  '시작날짜    
    txtTo_dt            = FilterVar(lgKeyStream(2), "''", "S")  '끝날짜 
    txtOcpt_type        = FilterVar(UCase(lgKeyStream(3)),"'%'", "S") '직종 
    txtSect_cd          = FilterVar(UCase(lgKeyStream(4)),"'%'", "S") '근무구역 
    txtFr_internal_cd   = FilterVar(lgKeyStream(5), "''", "S")  '시작부서 
    txtTo_internal_cd   = FilterVar(lgKeyStream(6), "''", "S")  '끝부서  
    rbo_sort            = lgKeyStream(7)                       '조회구분  
    txtPayCd            = FilterVar(UCase(lgKeyStream(8)),"'%'", "S") '지급구분 
    txtAllow_cd         = FilterVar(UCase(lgKeyStream(9)),"'%'", "S") '수당코드         
    '----------------TAB2------------------------------------------------------------
    txtFrom_dt2         = FilterVar(lgKeyStream(1), "''", "S")  '시작날짜  
    txtTo_dt2           = FilterVar(lgKeyStream(2), "''", "S")  '끝날짜 
    txtOcpt_type2       = FilterVar(lgKeyStream(3),"'%'", "S") '직종 
    txtSect_cd2         = FilterVar(lgKeyStream(4),"'%'", "S") '근무구역 
    txtFr_internal_cd2  = FilterVar(lgKeyStream(5), "''", "S")  '시작부서 
    txtTo_internal_cd2  = FilterVar(lgKeyStream(6), "''", "S")  '끝부서  
    rbo_sort2           = lgKeyStream(7)                       '조회구분  
    txtsub_type2        = FilterVar(lgKeyStream(8),"'%'", "S") '공제구분   
    txtsub_cd2          = FilterVar(lgKeyStream(9),"'%'", "S") '공제코드            
    
    If lgKeyStream(0)="1" AND rbo_sort="1" Then               '----수당변동사원(전체조회)
        FrSQL = lgKeyStream(5)       '  internal_cd = min
        ToSQL = lgKeyStream(6)       '  internal_cd = max
        
        strWhere = txtAllow_cd 
        strWhere = strWhere & " AND b.CODE_TYPE = " & FilterVar("1", "''", "S") & " "
        
        strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") >= " & FilterVar(lgKeyStream(5), "''", "S")
        strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") <= " & FilterVar(lgKeyStream(6), "''", "S")
        
        strWhere = strWhere & " AND a.OCPT_TYPE LIKE " & txtOcpt_type   '직종 
        strWhere = strWhere & " AND a.SECT_CD LIKE " & txtSect_cd      '근무구역 
        strWhere = strWhere & " AND a.EMP_NO = T1.EMP_NO "
        strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") LIKE  " & FilterVar(lgKeyStream(10) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'
     ElseIf lgkeyStream(0)="1" AND rbo_sort="2" Then       '----수당변동사원(변동사원조회)
            FrSQL = lgKeyStream(5)       '  internal_cd = min
            ToSQL = lgKeyStream(6)       '  internal_cd = max    
            
            strWhere = txtAllow_cd 
            strWhere = strWhere & " AND b.CODE_TYPE = " & FilterVar("1", "''", "S") & " "
            
            strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") >= " & FilterVar(lgKeyStream(5), "''", "S")
            strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") <= " & FilterVar(lgKeyStream(6), "''", "S")
            
            strWhere = strWhere & " AND a.OCPT_TYPE LIKE " & txtOcpt_type   '직종 
            strWhere = strWhere & " AND a.SECT_CD LIKE " & txtSect_cd      '근무구역 
            strWhere = strWhere & " AND (ISNULL(T2.ALLOW,0) - ISNULL(T1.ALLOW,0)) <> 0 "
            strWhere = strWhere & " AND a.EMP_NO = T1.EMP_NO "
            strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") LIKE  " & FilterVar(lgKeyStream(10) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'
     ElseIf lgkeyStream(0)="2" AND rbo_sort="1" Then       '----공제변동사원(전체조회)
            FrSQL2 = lgKeyStream(5)       '  internal_cd = min
            ToSQL2 = lgKeyStream(6)       '  internal_cd = max
            
            strWhere = txtsub_cd2 
            strWhere = strWhere & " AND b.CODE_TYPE = " & FilterVar("2", "''", "S") & ""
            
            strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") >= " & FilterVar(lgKeyStream(5), "''", "S")
            strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") <= " & FilterVar(lgKeyStream(6), "''", "S") 
            
            strWhere = strWhere & " AND a.OCPT_TYPE LIKE " & txtOcpt_type   '직종 
            strWhere = strWhere & " AND a.SECT_CD LIKE " & txtSect_cd      '근무구역 
            strWhere = strWhere & " AND a.EMP_NO = T1.EMP_NO "
            strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") LIKE  " & FilterVar(lgKeyStream(10) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'
     ElseIf lgkeyStream(0)="2" AND rbo_sort="2" Then       '----공제변동사원(변동사원조회)
            FrSQL2 = lgKeyStream(5)       '  internal_cd = min
            ToSQL2 = lgKeyStream(6)       '  internal_cd = max
            
            strWhere = txtsub_cd2 
            strWhere = strWhere & " AND b.CODE_TYPE = " & FilterVar("2", "''", "S") & ""
            
            strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") >= " & FilterVar(lgKeyStream(5), "''", "S")
            strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") <= " & FilterVar(lgKeyStream(6), "''", "S")
            
            strWhere = strWhere & " AND a.OCPT_TYPE LIKE " & txtOcpt_type   '직종 
            strWhere = strWhere & " AND a.SECT_CD LIKE " & txtSect_cd      '근무구역 
            strWhere = strWhere & " AND (ISNULL(T2.SUB_AMT,0) - ISNULL(T1.SUB_AMT,0)) <> 0 "
            strWhere = strWhere & " AND a.EMP_NO = T1.EMP_NO "
            strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(11),"'%'", "S") & ") LIKE  " & FilterVar(lgKeyStream(10) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'
    End If
   
    Call SubMakeSQLStatements("MR",strWhere,"X","=")                      '☆ : Make sql statements

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
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_GRD1"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ROLL_PSTN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ROLE"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BEFORE_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("CURRENT_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DIFF_AMT"), ggAmtOfMoney.DecPoint,0)
            
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgObjConn.Execute lgStrSQL,,adCmdText
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgObjConn.Execute lgStrSQL,,adCmdText
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub
'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgObjConn.Execute lgStrSQL,,adCmdText
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

                    If lgkeyStream(0)="1" Then  
                        
                        lgStrSQL = " SELECT "
                        lgStrSQL = lgStrSQL & " A.EMP_NO  EMP_NO, A.NAME  NAME, A.DEPT_NM  DEPT_NM, "
                        lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",A.PAY_GRD1) PAY_GRD1, "
                        lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",A.ROLL_PSTN) ROLL_PSTN, "
                        lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0026", "''", "S") & ",A.ROLE_CD) ROLE, "
                        lgStrSQL = lgStrSQL & " ISNULL(T1.ALLOW,0) BEFORE_AMT, ISNULL(T2.ALLOW,0) CURRENT_AMT, "
                        lgStrSQL = lgStrSQL & " ISNULL(T2.ALLOW,0) - ISNULL(T1.ALLOW,0) DIFF_AMT "
                        lgStrSQL = lgStrSQL & " FROM HAA010T A, HDA010T B, "
                        lgStrSQL = lgStrSQL & " (SELECT EMP_NO , ALLOW_CD , ALLOW FROM HDF040T WHERE PAY_YYMM = " & FilterVar(lgKeyStream(1), "''", "S")  
                        lgStrSQL = lgStrSQL & " AND PROV_TYPE  = " & FilterVar(lgKeyStream(8),"'%'", "S") 
                        lgStrSQL = lgStrSQL & " AND ALLOW_CD  = " & FilterVar(lgKeyStream(9),"'%'", "S")
                        lgStrSQL = lgStrSQL & " ) AS T1 FULL OUTER JOIN "
                        lgStrSQL = lgStrSQL & " (SELECT EMP_NO, ALLOW_CD, ALLOW FROM HDF040T WHERE PAY_YYMM = " & FilterVar(lgKeyStream(2), "''", "S") 
                        lgStrSQL = lgStrSQL & " AND PROV_TYPE = " & FilterVar(lgKeyStream(8),"'%'", "S")
                        lgStrSQL = lgStrSQL & " AND ALLOW_CD  = " & FilterVar(lgKeyStream(9),"'%'", "S")
                        lgStrSQL = lgStrSQL & " ) AS T2 ON (T1.EMP_NO = T2.EMP_NO) "
                        lgStrSQL = lgStrSQL & " Where b.ALLOW_CD " & pComp & pCode
                        lgStrSQL = lgStrSQL & " Order by A.INTERNAL_CD, A.EMP_NO"
             
                    ElseIf lgkeyStream(0)="2" Then 
                        
                        lgStrSQL = " SELECT "
                        lgStrSQL = lgStrSQL & " A.EMP_NO  EMP_NO, A.NAME  NAME, A.DEPT_NM  DEPT_NM, "
                        lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",A.PAY_GRD1) PAY_GRD1, "
                        lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",A.ROLL_PSTN) ROLL_PSTN, "
                        lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0026", "''", "S") & ",A.ROLE_CD) ROLE, "
                        lgStrSQL = lgStrSQL & " ISNULL(T1.SUB_AMT,0) BEFORE_AMT, ISNULL(T2.SUB_AMT,0) CURRENT_AMT, "
                        lgStrSQL = lgStrSQL & " ISNULL(T2.SUB_AMT,0) - ISNULL(T1.SUB_AMT,0) DIFF_AMT "
                        lgStrSQL = lgStrSQL & " FROM HAA010T A, HDA010T B, "
                        lgStrSQL = lgStrSQL & " (SELECT EMP_NO, SUB_CD, SUB_AMT FROM HDF060T WHERE SUB_YYMM = " & FilterVar(lgKeyStream(1), "''", "S") 
                        lgStrSQL = lgStrSQL & " AND SUB_TYPE = " & FilterVar(lgKeyStream(8),"'%'", "S") '공제구분 
                        lgStrSQL = lgStrSQL & " AND SUB_CD = " & FilterVar(lgKeyStream(9),"'%'", "S")  '공제코드 
                        lgStrSQL = lgStrSQL & " ) AS T1 FULL OUTER JOIN "
                        lgStrSQL = lgStrSQL & " (SELECT EMP_NO, SUB_CD, SUB_AMT FROM HDF060T WHERE SUB_YYMM = " & FilterVar(lgKeyStream(2), "''", "S") 
                        lgStrSQL = lgStrSQL & " AND SUB_TYPE = " & FilterVar(lgKeyStream(8),"'%'", "S") 
                        lgStrSQL = lgStrSQL & " AND SUB_CD = " & FilterVar(lgKeyStream(9),"'%'", "S") 
                        lgStrSQL = lgStrSQL & " ) AS T2 ON (T1.EMP_NO = T2.EMP_NO) "
                        lgStrSQL = lgStrSQL & " Where b.ALLOW_CD " & pComp & pCode
                        lgStrSQL = lgStrSQL & " Order by A.INTERNAL_CD, A.EMP_NO"
      
                   End If                
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
                If  Trim("<%=lgKeyStream(0)%>") = "1" Then
                   .ggoSpread.Source     = .frm1.vspdData
                   .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                Else
                   .ggoSpread.Source     = .frm1.vspdData2
                   .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                End If                
                
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
