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
    Dim lgObjRs1
    Dim lgStrSQL1
    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")    

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    
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
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
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
    Dim strFr_internal_cd
    Dim strTo_internal_cd
    Dim strWhere
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    IF lgKeyStream(10) = "Del" THEN  ' 중복 제거 
        lgStrSQL = " DELETE FROM HDF050T   FROM HDF050T A, HDF020T B"
        lgStrSQL = lgStrSQL & " WHERE A.EMP_NO	= B.EMP_NO "
        lgStrSQL = lgStrSQL & " AND A.EMP_NO      LIKE " & FilterVar(lgKeyStream(0),"'%'", "S")
        lgStrSQL = lgStrSQL & " AND A.SUB_TYPE    LIKE " & FilterVar(lgKeyStream(1),"'%'", "S")
        lgStrSQL = lgStrSQL & " AND A.SUB_CD      LIKE " & FilterVar(lgKeyStream(2),"'%'", "S")

        lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	end if    
    
    
    if lgKeyStream(10) = "Add" THEN  ' 중복아닌 것만 추가 
		lgStrSQL = "SELECT A.NAME, A.EMP_NO, A.DEPT_CD, B.DEPT_NM , A.INTERNAL_CD, dbo.ufn_getCodeName(" & FilterVar("H0040", "''", "S") & ","
		lgStrSQL = lgStrSql & FilterVar(lgKeyStream(1), "''", "S")  & " ) Sub_type_NM, "
		lgStrSQL = lgStrSql & " dbo.ufn_H_GetCodeName(" & FilterVar("HDA010t", "''", "S") & "," & FilterVar(lgKeyStream(2), "''", "S") & ",'') SUB_CD_NM "
		lgStrSQL = lgStrSQL & " FROM  HDF020T A, HAA010T B "
		lgStrSQL = lgStrSQL & " WHERE A.EMP_NO = B.EMP_NO AND A.EMP_NO not in "
		lgStrSQL = lgStrSQL & " (select A.EMP_NO FROM HDF050T A, HDF020T B"
        lgStrSQL = lgStrSQL & "  WHERE A.EMP_NO	= B.EMP_NO "
        lgStrSQL = lgStrSQL & "        AND A.EMP_NO      LIKE " & FilterVar(lgKeyStream(0),"'%'", "S")
        lgStrSQL = lgStrSQL & "        AND A.SUB_TYPE    LIKE " & FilterVar(lgKeyStream(1),"'%'", "S")
        lgStrSQL = lgStrSQL & "        AND A.SUB_CD      LIKE " & FilterVar(lgKeyStream(2),"'%'", "S") & ") "
		lgStrSQL = lgStrSQL & " AND A.PROV_TYPE = " & FilterVar("Y", "''", "S") & "  "
		lgStrSQL = lgStrSQL & " AND A.RETIRE_DT IS NULL AND A.EMP_NO LIKE " & FilterVar(lgKeyStream(0),"'%'", "S")
		lgStrSQL = lgStrSQL & " AND A.INTERNAL_CD BETWEEN  " & FilterVar(lgKeyStream(7), "''", "S") & " AND  " & FilterVar(lgKeyStream(8), "''", "S") & ""  'internal_cd = max
		lgStrSQL = lgStrSQL & " AND A.INTERNAL_CD LIKE  " & FilterVar(lgKeyStream(9) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'
		lgStrSQL = lgStrSQL & " ORDER BY A.DEPT_CD , B.EMP_NO "
    
    elseif  lgKeyStream(10) = "Normal" OR lgKeyStream(10) = "Del"THEN  ' 중복없음 
		lgStrSQL = "SELECT A.NAME, A.EMP_NO, A.DEPT_CD, B.DEPT_NM , A.INTERNAL_CD, dbo.ufn_getCodeName(" & FilterVar("H0040", "''", "S") & ","
		lgStrSQL = lgStrSql & FilterVar(lgKeyStream(1), "''", "S")  & " ) Sub_type_NM, "
		lgStrSQL = lgStrSql & " dbo.ufn_H_GetCodeName(" & FilterVar("HDA010t", "''", "S") & "," & FilterVar(lgKeyStream(2), "''", "S") & ", '') SUB_CD_NM "
		lgStrSQL = lgStrSQL & " FROM  HDF020T A, HAA010T B "
		lgStrSQL = lgStrSQL & " WHERE A.EMP_NO = B.EMP_NO AND A.PROV_TYPE = " & FilterVar("Y", "''", "S") & "  "
		lgStrSQL = lgStrSQL & " AND A.RETIRE_DT IS NULL AND A.EMP_NO LIKE " & FilterVar(lgKeyStream(0),"'%'", "S")
		lgStrSQL = lgStrSQL & " AND A.INTERNAL_CD BETWEEN  " & FilterVar(lgKeyStream(7), "''", "S") & " AND  " & FilterVar(lgKeyStream(8), "''", "S") & ""  'internal_cd = max
		lgStrSQL = lgStrSQL & " AND A.INTERNAL_CD LIKE  " & FilterVar(lgKeyStream(9) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'
		lgStrSQL = lgStrSQL & " ORDER BY A.DEPT_CD , B.EMP_NO "
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
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM") )  
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgKeyStream(1))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgKeyStream(2))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_CD_NM"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & lgKeyStream(13)
            lgstrData = lgstrData & Chr(11) & lgKeyStream(11)
            lgstrData = lgstrData & Chr(11) & lgKeyStream(12)  'lgObjRs("REVOKE_YYMM")
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
