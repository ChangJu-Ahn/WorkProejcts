<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")

    Dim lgObjRs1
    Dim lgStrSQL1

  
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    
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
    Dim iKey1
    Dim strWhere
    Dim IntRetCD
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If lgKeyStream(0) = "" then
        lgKeyStream(0) = "%"
    End if
    
    If Trim(lgKeyStream(7)) = "Del" then             '재생성시 
        lgStrSQL = "DELETE  HDF030T "
        lgStrSQL = lgStrSQL & " WHERE        "
        lgStrSQL = lgStrSQL & " EMP_NO   IN (SELECT emp_no FROM hdf020t WHERE emp_no LIKE " & FilterVar(lgKeyStream(0), "''", "S")
        
        lgStrSQL = lgStrSQL & " AND dbo.ufn_H_get_internal_cd(emp_no," & FilterVar(lgKeyStream(9),"'%'", "S") & ") >=  " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & " AND dbo.ufn_H_get_internal_cd(emp_no," & FilterVar(lgKeyStream(9),"'%'", "S") & ") <=  " & FilterVar(lgKeyStream(2), "''", "S")
        
        lgStrSQL = lgStrSQL & " ) AND allow_cd LIKE  " & FilterVar(lgKeyStream(3), "''", "S")
        
        lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
        Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    End if
	if lgKeyStream(7) = "Add" THEN  ' 중복아닌 것만 추가 
		If lgKeyStream(0) = "" then
			strWhere = "" & FilterVar("%", "''", "S") & ""
		Else
			strWhere = FilterVar(lgKeyStream(0), "''", "S")
		End if
		
		strWhere = strWhere & " AND a.EMP_NO not in ( select EMP_NO from hdf030t"
		strWhere = strWhere & " WHERE        "
        strWhere = strWhere & " EMP_NO  IN (SELECT emp_no FROM hdf020t WHERE emp_no LIKE " & FilterVar(lgKeyStream(0), "''", "S")


        strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(emp_no," & FilterVar(lgKeyStream(9),"'%'", "S") & ") >=  " & FilterVar(lgKeyStream(1), "''", "S")
        strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(emp_no," & FilterVar(lgKeyStream(9),"'%'", "S") & ") <=  " & FilterVar(lgKeyStream(2), "''", "S")


        strWhere = strWhere & " ) AND allow_cd LIKE  " & FilterVar(lgKeyStream(3), "''", "S") & ") "

		strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(9),"'%'", "S") & ") >= " & FilterVar(lgKeyStream(1), "''", "S") 
		strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(9),"'%'", "S") & ") <= " & FilterVar(lgKeyStream(2), "''", "S")

		lgStrSQL = "SELECT a.NAME, a.EMP_NO, a.DEPT_CD, b.DEPT_NM " 
		lgStrSQL = lgStrSQL & "  FROM HDF020T a, HAA010T b "
		lgStrSQL = lgStrSQL & " WHERE a.EMP_NO = b.EMP_NO AND a.PROV_TYPE = " & FilterVar("Y", "''", "S") & "  AND a.RETIRE_DT IS NULL "
		lgStrSQL = lgStrSQL & "   AND a.EMP_NO LIKE  " & strWhere
	
	elseif lgKeyStream(7) = "Normal" OR lgKeyStream(7) = "Del"THEN  ' 중복없음 
		If lgKeyStream(0) = "" then
			strWhere = "" & FilterVar("%", "''", "S") & ""
		Else
			strWhere = FilterVar(lgKeyStream(0), "''", "S")
		End if
		
		strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(9),"'%'", "S") & ") >= " & FilterVar(lgKeyStream(1), "''", "S") 
		strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(9),"'%'", "S") & ") <= " & FilterVar(lgKeyStream(2), "''", "S")

		lgStrSQL = "SELECT a.NAME, a.EMP_NO, a.DEPT_CD, b.DEPT_NM " 
		lgStrSQL = lgStrSQL & "  FROM HDF020T a, HAA010T b "
		lgStrSQL = lgStrSQL & " WHERE a.EMP_NO = b.EMP_NO AND a.PROV_TYPE = " & FilterVar("Y", "''", "S") & "  AND a.RETIRE_DT IS NULL "
		lgStrSQL = lgStrSQL & "   AND a.EMP_NO LIKE  " & strWhere
	end if 
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        lgstrData = ""

        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & lgKeyStream(3)
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & lgKeyStream(8)
                
            lgstrData = lgstrData & Chr(11) & lgKeyStream(4)
            lgstrData = lgstrData & Chr(11) & lgKeyStream(5)
            lgstrData = lgstrData & Chr(11) & lgKeyStream(6)

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
	    	    
        Loop 
    End If
    
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
    Call SubCloseRs(lgObjRs1)                                                         '☜: Release RecordSSet

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
       Case "<%=UID_M0001%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBAutoQueryOk
	         End with
          End If   
    End Select    
    
       
</Script>	
