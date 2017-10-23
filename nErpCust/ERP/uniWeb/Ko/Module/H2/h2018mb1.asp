<%@ LANGUAGE=VBSCript%>
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
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")

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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    if  Trim(lgKeyStream(0)) <> "" then
         iKey1 = " AND EMP_NO LIKE  " & FilterVar(Trim(lgKeyStream(0)) & "%", "''", "S")  
    end if

    If lgKeyStream(1) <> "" AND lgKeyStream(2) <> "" Then
        iKey1 = iKey1 & " AND ISSUE_DT >= " & FilterVar(UNIConvDate(lgKeyStream(1)), "''", "S")
        iKey1 = iKey1 & " AND ISSUE_DT <= " & FilterVar(UNIConvDate(lgKeyStream(2)), "''", "S")
    End If

    if  Trim(lgKeyStream(3)) <> "" and  Trim(lgKeyStream(3)) <> "4"  then
         iKey1 =  iKey1 &" AND PROOF_TYPE = " & FilterVar(Trim(lgKeyStream(3)), "''", "S")  
    end if

    Call SubMakeSQLStatements("MR",iKey1,"X","")                                 '☆ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROOF_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_NO"))
            
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("ISSUE_DT"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PRINT_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROOF_USE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PRINT_EMP_NO"))
            
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
                       lgStrSQL = lgStrSQL & "	EMP_NO, "
                       lgStrSQL = lgStrSQL & "	dbo.ufn_H_GetEmpName(EMP_NO) NAME, "
                       lgStrSQL = lgStrSQL & "	case when PROOF_TYPE='1' then '재직' when  PROOF_TYPE='2' then '경력' when  PROOF_TYPE='3' then '퇴직' end PROOF_TYPE, "
                       lgStrSQL = lgStrSQL & "	ISSUE_NO + '-' + CONVERT(VARCHAR(20),ISSUE_NO1) ISSUE_NO, "
                       lgStrSQL = lgStrSQL & "	ISSUE_DT, "                       
                       lgStrSQL = lgStrSQL & "	PRINT_DT, "
                       lgStrSQL = lgStrSQL & "	PROOF_USE, "
                       lgStrSQL = lgStrSQL & "	case when dbo.ufn_H_GetEmpName(PRINT_EMP_NO) = '' then PRINT_EMP_NO else dbo.ufn_H_GetEmpName(PRINT_EMP_NO) end PRINT_EMP_NO "
                       lgStrSQL = lgStrSQL & " FROM  HAA170T "
                       lgStrSQL = lgStrSQL & " WHERE 1=1 " & pCode
                    
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
    End Select    
    
       
</Script>	
