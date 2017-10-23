<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
    Dim lgObjRs1
    Dim lgStrSQL1
    Dim strWk_type
    Dim strWk_type_nm
    Dim strHoli_type
    Dim strHoli_type_nm
    
    call LoadBasisGlobalInf()
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Query
             Call SubBizQuery1()
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
	Dim strWhere
	Dim strWhere1
	Dim strEmp_no, strChang_dt

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	strEmp_no = FilterVar(lgKeyStream(1), "''", "S")
	
    If lgKeyStream(1) = "" then
    Else
       strWhere = strEmp_no
    End if
    
    If lgKeyStream(0) = "" then
       strWhere = strWhere & " AND chang_dt = " & "" & FilterVar("%", "''", "S") & ""
    Else
       strWhere = strWhere & " AND chang_dt =(SELECT max(chang_dt) from hca040t where emp_no= " & FilterVar(lgKeyStream(1), "''", "S")
       strWhere = strWhere & " AND chang_dt <= " & FilterVar(UNIConvDateCompanyToDB(lgKeyStream(0),NULL), "''", "S") & ")" 
    End if 

    lgStrSQL = "Select wk_type, minor_nm   "                   '근무조(wk_type)을 구한다.
    lgStrSQL = lgStrSQL & " From  HCA040T a, B_MINOR b "
    lgStrSQL = lgStrSQL & " WHERE a.emp_no = " & strWhere
    lgStrSQL = lgStrSQL & "  AND  a.wk_type = b.minor_cd "
    lgStrSQL = lgStrSQL & "  AND  b.major_cd = " & FilterVar("H0047", "''", "S") & " "
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        strWk_type = "0"
        strWk_type_nm = "공통조"
    Else
        strWk_type = ConvSPChars(lgObjRs("wk_type"))
        strWk_type_nm = ConvSPChars(lgObjRs("minor_nm"))
    End if
    
    strWhere1 = FilterVar(UNIConvDateCompanyToDB(lgKeyStream(0),NULL),"NULL","S")
    strWhere1 = strWhere1 & " AND wk_type =  " & FilterVar(strWk_type , "''", "S") & "" 
    strWhere1 = strWhere1 & " AND org_cd =( "
    strWhere1 = strWhere1 & " SELECT B.BIZ_AREA_CD   FROM B_ACCT_DEPT A, B_COST_CENTER B "
    strWhere1 = strWhere1 & " WHERE A.COST_CD = B.COST_CD "
    strWhere1 = strWhere1 & " AND A.DEPT_CD = (SELECT DEPT_CD   FROM HAA010T  WHERE EMP_NO = " & FilterVar(lgKeyStream(1), "''", "S")
    strWhere1 = strWhere1 & " AND A.ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= GETDATE()) "

    lgStrSQL = "Select holi_type, minor_nm   "                  '휴일(holi_type)을 구한다.....
    lgStrSQL = lgStrSQL & " From  HCA020T a, B_MINOR b "
    lgStrSQL = lgStrSQL & " WHERE a.date = " & strWhere1 & "))"
    lgStrSQL = lgStrSQL & "  AND  a.holi_type = b.minor_cd "
    lgStrSQL = lgStrSQL & "  AND  b.major_cd = " & FilterVar("H0049", "''", "S") & " "

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        strHoli_type = "D"
        strHoli_type_nm = "평일"
    Else
        strHoli_type = ConvSPChars(lgObjRs("holi_type"))
        strHoli_type_nm = ConvSPChars(lgObjRs("minor_nm"))
    end if
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

Sub SubBizQuery1()
    Dim iDx,iSelCount
    Dim iLoopMax
    Dim IntRetCD
    Dim strWhere
    Dim strAttend_dt
    dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   
    strAttend_dt = FilterVar(UNIConvDateCompanyToDB(lgKeyStream(0),NULL),"NULL","S")
    If lgKeyStream(0) = "" then
    Else
       strWhere = strAttend_dt & ")"
       strWhere = strWhere & " AND entr_dt <= " & strAttend_dt
    End if
    
    If lgKeyStream(1) = "" then
       strWhere = strWhere & " AND a.emp_no LIKE " & "" & FilterVar("%", "''", "S") & ""
    Else
       strWhere = strWhere & " AND a.emp_no LIKE " & FilterVar(lgKeyStream(1), "''", "S")
    End if 
    
    If lgKeyStream(7) = "" then
        strWhere = strWhere & " AND internal_cd LIKE  " & FilterVar(Trim(lgKeyStream(2)) & "%", "''", "S") & ""
    Else
        strWhere = strWhere & " AND internal_cd = " & FilterVar(lgKeyStream(2), "''", "S")
    End if 

    lgStrSQL = "SELECT  a.name, a.emp_no, dbo.ufn_H_get_dept_cd(a.EMP_NO,"& strAttend_dt & ") dept_cd,"
    lgStrSQL = lgStrSQL & " dbo.ufn_H_get_internal_cd(a.EMP_NO,"& strAttend_dt & ") internal_cd, "
    lgStrSQL = lgStrSQL & " dbo.ufn_GetDeptName( dbo.ufn_H_get_dept_cd(a.EMP_NO,"& strAttend_dt & "),"& strAttend_dt &" ) DEPT_CD_NM, " 
    
    lgStrSQL = lgStrSQL & " b.wk_type wk_type, dbo.ufn_GetCodeName(" & FilterVar("H0047", "''", "S") & ",b.wk_type) wk_type_nm, "
    lgStrSQL = lgStrSQL & " c.holi_type holi_type, dbo.ufn_GetCodeName(" & FilterVar("H0049", "''", "S") & ",c.holi_type) holi_type_nm "
    lgStrSQL = lgStrSQL & " FROM  HAA010T a, hca040t b, hca020t c "
    lgStrSQL = lgStrSQL & " WHERE (retire_dt IS NULL OR retire_dt >= " & strWhere
    lgStrSQL = lgStrSQL & "   AND a.emp_no = b.emp_no "
    lgStrSQL = lgStrSQL & "   AND b.chang_dt = (SELECT max(chang_dt) FROM hca040t WHERE emp_no = a.emp_no AND chang_dt <= " & strAttend_dt & ")"
    lgStrSQL = lgStrSQL & "   AND b.wk_type = c.wk_type "
    lgStrSQL = lgStrSQL & "   AND c.date = " & strAttend_dt
    lgStrSQL = lgStrSQL & "   AND c.org_cd = (SELECT y.biz_area_cd from b_acct_dept x, b_cost_center y "
    lgStrSQL = lgStrSQL &          " WHERE x.dept_cd = a.dept_cd "
    lgStrSQL = lgStrSQL &          "   AND x.cost_cd = y.cost_cd "
    lgStrSQL = lgStrSQL &          "   AND x.org_change_dt = (SELECT max(org_change_dt) FROM b_acct_dept WHERE org_change_dt <= " & strAttend_dt & "))"

    If 	FncOpenRs("R",lgObjConn,lgObjRs1,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        lgstrData = ""
       
        Do While Not lgObjRs1.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs1("EMP_NO"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs1("NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs1("DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs1("DEPT_CD_NM"))
            lgstrData = lgstrData & Chr(11) & lgKeyStream(0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgKeyStream(3)))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgKeyStream(4)))
            lgstrData = lgstrData & Chr(11) & lgKeyStream(0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgKeyStream(5)))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgKeyStream(6)))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs1("wk_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs1("wk_type_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs1("holi_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs1("holi_type_nm"))

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs1("INTERNAL_CD"))

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs1.MoveNext
        Loop 
    End If
	Call SubHandleError("MR",lgObjConn,lgObjRs1,Err)
    Call SubCloseRs(lgObjRs1)                                                          '☜: Release RecordSSet

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
                .frm1.vspdData.col = 12
                .frm1.vspdData.text = "<%=strWk_type%>"

                .frm1.vspdData.col = 13
                .frm1.vspdData.text = "<%=strWk_type_nm%>"

                .frm1.vspdData.col = 14
                .frm1.vspdData.text = "<%=strHoli_type%>"

                .frm1.vspdData.col = 15
                .frm1.vspdData.text = "<%=strHoli_type_nm%>"
	         End with
          End If
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBAutoQueryOk
	         End with
          End If   
    End Select    
    
       
</Script>	
