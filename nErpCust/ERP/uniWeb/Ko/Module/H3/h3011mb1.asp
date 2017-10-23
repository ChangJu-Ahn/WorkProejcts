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
             Call SubBizQueryMulti()
     End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
 
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1, baseDt

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	baseDt = FilterVar(UNIConvDate(lgKeyStream(4)), "''", "S")

    iKey1 = " AND A.EMP_NO LIKE  " & FilterVar(Trim(lgKeyStream(0)) & "%", "''", "S") & ""
    if  Trim(lgKeyStream(2)) <> "" then
        iKey1 = iKey1 & " AND EDU_CD = " & FilterVar(lgKeyStream(2), "''", "S")
    end if

    if  Trim(lgKeyStream(3)) <> "" then
        iKey1 = iKey1 & " AND EDU_OFFICE = " & FilterVar(lgKeyStream(3), "''", "S")
    end if

    If lgKeyStream(4) <> "" AND lgKeyStream(5) <> "" Then
        iKey1 = iKey1 & " AND EDU_START_DT >= " & FilterVar(UNIConvDate(lgKeyStream(4)), "''", "S")
        iKey1 = iKey1 & " AND EDU_END_DT <= " & FilterVar(UNIConvDate(lgKeyStream(5)), "''", "S")
    End If

    If lgKeyStream(6) <> ""   Then
        iKey1 = iKey1 &  " AND dbo.ufn_H_get_dept_cd (a.EMP_NO ," & baseDt & ") =" & FilterVar(lgKeyStream(6), "''", "S")
    End If
 	
    Call SubMakeSQLStatements("MR",iKey1,baseDt,"")                                 '☆ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EDU_nm"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("EDU_START_DT"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("EDU_END_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EDU_OFFICE_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EDU_NAT_nm"))
         
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EDU_CONT"))

            If  lgObjRs("edu_type") = "1" Then
                lgstrData = lgstrData & Chr(11) & "필수"
            Else
                lgstrData = lgstrData & Chr(11) & "선택"
            End If
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("EDU_SCORE"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("END_DT"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("EDU_FEE"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FEE_TYPE"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("REPAY_AMT"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REPORT_TYPE"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ADD_POINT"), ggAmtOfMoney.DecPoint, 0)

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
                       lgStrSQL = lgStrSQL & "	dbo.ufn_H_get_dept_cd (a.EMP_NO ," & pCode1 & ") dept_cd,"
					   lgStrSQL = lgStrSQL & "	dbo.ufn_GetDeptName(dbo.ufn_H_get_dept_cd (a.EMP_NO , " & pCode1 & " )," & pCode1 & ") dept_nm, "
                       lgStrSQL = lgStrSQL & "	A.EMP_NO, "
                       lgStrSQL = lgStrSQL & "	B.NAME, "
                       lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0033", "''", "S") & ",EDU_CD) EDU_nm, "
                       lgStrSQL = lgStrSQL & "	EDU_START_DT, "
                       lgStrSQL = lgStrSQL & "	EDU_END_DT, "                       
                       lgStrSQL = lgStrSQL & "	dbo.ufn_H_GetCodeName(" & FilterVar("B_COUNTRY", "''", "S") & ",EDU_NAT,'') EDU_NAT_nm, "
                       lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0037", "''", "S") & ",EDU_OFFICE) EDU_OFFICE_nm, "
                       lgStrSQL = lgStrSQL & "	EDU_CONT, "
                       lgStrSQL = lgStrSQL & "	EDU_TYPE, "
                       lgStrSQL = lgStrSQL & "	EDU_FEE, "
                       lgStrSQL = lgStrSQL & "	EDU_SCORE, "
                       lgStrSQL = lgStrSQL & "	ADD_POINT, "
                       lgStrSQL = lgStrSQL & "	END_DT, "
                       lgStrSQL = lgStrSQL & "	FEE_TYPE, "
                       lgStrSQL = lgStrSQL & "	REPORT_TYPE, "
                       lgStrSQL = lgStrSQL & "	REPAY_AMT "
                       lgStrSQL = lgStrSQL & " FROM  HBA030T A, HAA010T B "
                       lgStrSQL = lgStrSQL & " WHERE A.EMP_NO = B.EMP_NO " & pCode
                       lgStrSQL = lgStrSQL & " ORDER BY dbo.ufn_H_get_dept_cd (a.EMP_NO ," & pCode1 & ") ,a.EMP_NO "
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
