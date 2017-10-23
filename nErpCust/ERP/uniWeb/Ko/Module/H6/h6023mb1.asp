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
    Dim strFrom, strWhere
	Const C_SHEETMAXROWS_D = 100

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("Q", "H","NOCOOKIE","MB")

    Dim lgGetSvrDateTime, lgStrPrevKey    

    lgGetSvrDateTime = GetSvrDateTime

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey      = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

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
    
  	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
  	Dim txtWhere, strTrans_dt
  	
  	txtWhere = " PROV_DT = " & FilterVar(lgKeyStream(0), "''", "S") & " AND PROV_TYPE = " & FilterVar(lgKeyStream(1), "''", "S")
  	
	Call CommonQueryRs(" DISTINCT TRANS_DT "," HDF100T ", txtWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	strTrans_dt = UniDateClientFormat(Replace(lgF0, Chr(11), ""))

    %>
        <Script Language=vbscript>
            With Parent.Frm1
                 .txtTrans_dt.Text = "<%=strTrans_dt%>"
            End With  
        </Script>
    <%  

    Call SubBizQueryMulti()
End Sub    
	    
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	strWhere = FilterVar(lgKeyStream(0), "''", "S")
	strWhere = strWhere & " AND PROV_TYPE = " & FilterVar(lgKeyStream(1), "''", "S")

	If trim(lgKeyStream(2)) <> "" Then
		strWhere = strWhere & " AND BIZ_AREA_CD = " & FilterVar(trim(lgKeyStream(2)), "''", "S")
	End If
	
	If trim(lgKeyStream(3)) <> "" Then
		strWhere = strWhere & " AND ACCT_CD = " & FilterVar(trim(lgKeyStream(3)), "''", "S")
	End If
	
	If trim(lgKeyStream(4)) <> "" Then		
		strWhere = strWhere & " AND CODE_TYPE = " & FilterVar(trim(lgKeyStream(4)), "''", "S")
	End If 
	
	If trim(lgKeyStream(5)) <> "" Then
		strWhere = strWhere & " AND ALLOW_CD = " & FilterVar(trim(lgKeyStream(5)), "''", "S")
	End If 

    strFrom =     "SELECT     dbo.ufn_H_GetCodeName('B_BIZ_AREA',BIZ_AREA_CD,'') biz_area_cd" 
    strFrom = strFrom & "	, dbo.ufn_H_GetCodeName('A_ACCT',ACCT_CD,'') accnt "
    strFrom = strFrom & "	, dbo.ufn_GetDeptName(DEPT_CD,getdate()) dept_nm"
    strFrom = strFrom & "	, case when CODE_TYPE ='1' then AMT else 0 end allow"
    strFrom = strFrom & "	, case when CODE_TYPE ='2' then AMT else 0 end ded"
    strFrom = strFrom & " FROM HDF110T  "
    strFrom = strFrom & " WHERE PROV_DT = " & strWhere

    If lgKeyStream(6) = "1" Then       '계정별합 
        Call SubMakeSQLStatements("MR","1A")
    ElseIf lgKeyStream(6) = "2" Then       '부서별 
        Call SubMakeSQLStatements("MR","1B")
    ElseIf lgKeyStream(6) = "3" Then       '개인별 
        Call SubMakeSQLStatements("MR","1C")
    End If

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs, C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF

            If lgKeyStream(6) = "1" Then       '계정별 
                lgstrData = lgstrData & Chr(11) & lgObjRs("accnt")

                If lgObjRs("accnt") = "총계" Then
                    lgstrData = lgstrData & Chr(11) & ""
                Else
                    lgstrData = lgstrData & Chr(11) & lgObjRs("biz_area_cd")
                End If
            ElseIf lgKeyStream(6) = "2" Then       '부서별 
                lgstrData = lgstrData & Chr(11) & lgObjRs("biz_area_cd")
                If lgObjRs("biz_area_cd") = "총계" Then
                    lgstrData = lgstrData & Chr(11) & ""
                Else
                    lgstrData = lgstrData & Chr(11) & lgObjRs("accnt")
                End If
               
                If lgObjRs("accnt") = "합계" Then
                    lgstrData = lgstrData & Chr(11) & ""
                Else
                    lgstrData = lgstrData & Chr(11) & lgObjRs("dept_nm")
                End If
            ElseIf lgKeyStream(6) = "3" Then       '개인별 
                lgstrData = lgstrData & Chr(11) & lgObjRs("biz_area_cd")
                lgstrData = lgstrData & Chr(11) & lgObjRs("accnt")
                lgstrData = lgstrData & Chr(11) & lgObjRs("allow_cd")
                lgstrData = lgstrData & Chr(11) & lgObjRs("dept_nm")
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            End If

            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("allow_amt"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ded_amt"), ggAmtOfMoney.DecPoint, 0)
    
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

  '개인별이 아니고 총계일때 
            If lgKeyStream(6) <> "3" AND (lgObjRs("biz_area_cd") = "총계" OR lgObjRs("accnt") = "총계") Then 
                %>
                <Script Language=vbscript>
                      With Parent.Frm1
                        .txtDrLocAmt.Value = "<%=UNINumClientFormat(lgObjRs("allow_amt"), ggAmtOfMoney.DecPoint, 0)%>"
                        .txtCrLocAmt.Value = "<%=UNINumClientFormat(lgObjRs("ded_amt"), ggAmtOfMoney.DecPoint, 0)%>"
                      End With  
                </Script>
                <%  
            End If

		    lgObjRs.MoveNext

            iDx =  iDx + 1
 
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If 
        Loop 
    End If

    If lgKeyStream(6) = "3" Then       '개인별 
        Call SubMakeSQLStatements("SR","1C")
        
        If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
            lgStrPrevKey = ""
            Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
            Call SetErrorStatus()
        Else
            %>
            <Script Language=vbscript>
                  With Parent.Frm1
                    .txtDrLocAmt.Value = "<%=UNINumClientFormat(lgObjRs("allow_amt"), ggAmtOfMoney.DecPoint, 0)%>"
                    .txtCrLocAmt.Value = "<%=UNINumClientFormat(lgObjRs("ded_amt"), ggAmtOfMoney.DecPoint, 0)%>"
                  End With  
            </Script>
            <%  
        End If
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements'("MR",iKey1,"X",C_EQ), ("MR",iKey1,"X")
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode)
     Dim iSelCount
     Dim iSql_allow, iSql_ded, iSql_bonus

     iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

     Select Case  pDataType     
     Case "MR"
        Select Case  pCode 
        Case "1A"
           lgStrSQL =     "SELECT  TOP " & iSelCount  
           lgStrSQL = lgStrSQL & " CASE WHEN (GROUPING(accnt) = 1)			THEN '총계'  ELSE accnt			END AS accnt , "
           lgStrSQL = lgStrSQL & " CASE WHEN (GROUPING(biz_area_cd) = 1)	THEN '합계'  ELSE biz_area_cd	END AS biz_area_cd , "
           lgStrSQL = lgStrSQL & " sum(allow) allow_amt, sum(ded) ded_amt "
           lgStrSQL = lgStrSQL & " FROM ( " & strFrom & " ) X "
           lgStrSQL = lgStrSQL & " GROUP BY accnt, biz_area_cd WITH ROLLUP "
        Case "1B"
           lgStrSQL =     "SELECT  TOP " & iSelCount  
           lgStrSQL = lgStrSQL & " CASE WHEN (GROUPING(biz_area_cd) = 1)	THEN	'총계'  ELSE biz_area_cd	END AS biz_area_cd , "
           lgStrSQL = lgStrSQL & " CASE WHEN (GROUPING(accnt) = 1)			THEN	'합계'  ELSE accnt			END AS accnt , "
           lgStrSQL = lgStrSQL & " CASE WHEN (GROUPING(dept_nm) = 1)		THEN	'소계'  ELSE dept_nm		END AS dept_nm , "
           lgStrSQL = lgStrSQL & " sum(allow) allow_amt, sum(ded) ded_amt "
           lgStrSQL = lgStrSQL & " FROM ( " & strFrom & " ) X "
           lgStrSQL = lgStrSQL & " GROUP BY biz_area_cd, accnt, dept_nm WITH ROLLUP "

        Case "1C"	'개인별조회 
           lgStrSQL =     "SELECT  TOP " & iSelCount  
           lgStrSQL = lgStrSQL & "	  dbo.ufn_H_GetCodeName('B_BIZ_AREA',BIZ_AREA_CD,'') biz_area_cd" 
           lgStrSQL = lgStrSQL & "	, dbo.ufn_H_GetCodeName('A_ACCT',ACCT_CD,'') accnt "
           lgStrSQL = lgStrSQL & "	, dbo.ufn_H_GetCodeName('HDA010T',ALLOW_CD,'') allow_cd "
           lgStrSQL = lgStrSQL & "	, dbo.ufn_GetDeptName(DEPT_CD,getdate()) dept_nm"
           lgStrSQL = lgStrSQL & "	, EMP_NO"
           lgStrSQL = lgStrSQL & "	, dbo.ufn_H_GetEmpName(EMP_NO) name"
           lgStrSQL = lgStrSQL & "	, case when CODE_TYPE ='1' then AMT else 0 end allow_amt"
           lgStrSQL = lgStrSQL & "	, case when CODE_TYPE ='2' then AMT else 0 end ded_amt"
           lgStrSQL = lgStrSQL & " FROM HDF110T  "
           lgStrSQL = lgStrSQL & " WHERE PROV_DT = " & strWhere
           lgStrSQL = lgStrSQL & " ORDER BY BIZ_AREA_CD, ACCT_CD, ALLOW_CD, DEPT_CD, EMP_NO "
        End Select
        
     Case "SR"            '개인별일때 합계부 
        Select Case  pCode 
        Case "1C"
           lgStrSQL = " SELECT SUM(allow_amt) allow_amt, SUM(ded_amt) ded_amt "
           lgStrSQL = lgStrSQL & " FROM ( SELECT  case when CODE_TYPE ='1' then SUM(AMT) else 0 end allow_amt"                     
           lgStrSQL = lgStrSQL & "				, case when CODE_TYPE ='2' then SUM(AMT) else 0 end ded_amt"
           lgStrSQL = lgStrSQL & "		 FROM HDF110T  "
           lgStrSQL = lgStrSQL & "		WHERE PROV_DT = " & strWhere    
           lgStrSQL = lgStrSQL & "		GROUP BY CODE_TYPE "
           lgStrSQL = lgStrSQL & " ) X "
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"                               '☜ : Display data                
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .DBQueryOk        
             End with
          End If   
    End Select    
       
</Script>	

