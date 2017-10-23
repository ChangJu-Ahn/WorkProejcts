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
	Const C_SHEETMAXROWS_D = 100

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("Q", "H","NOCOOKIE","MB")

    Dim lgGetSvrDateTime, lgStrPrevKey    
    lgGetSvrDateTime = GetSvrDateTime

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")                                       '¢Ð: "P"(Prev search) "N"(Next search)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey      = UNICInt(Trim(Request("lgStrPrevKey")),0)                    '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQueryMulti()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection
	    
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1
    Dim iLoopMax

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
	if Trim(lgKeyStream(1)) = "" then
		iKey1 = " c.emp_no LIKE " & FilterVar("%", "''", "S") & ""
	else
		iKey1 = " c.emp_no  LIKE " & FilterVar(lgKeyStream(1),"'%'", "S")
	end if	

	if Trim(lgKeyStream(0)) = "" then		
		iKey1 = iKey1 & " AND y.biz_area_cd LIKE " & FilterVar("%", "''", "S") & " AND"	
	else
		iKey1 = iKey1 & " AND y.biz_area_cd = " & FilterVar(lgKeyStream(0), "''", "S") & " AND"
	end if
    
    Call SubMakeSQLStatements("MR",iKey1)

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '¢Ð : No data is found.
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs, C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx = 1

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("emp_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("res_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bank"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bank_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bankmaster"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bank_accnt"))
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
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements'("MR",iKey1,"X",C_EQ), ("MR",iKey1,"X",C_LIKE)
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode)
     Dim iSelCount

     iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

     Select Case  pDataType 
        Case "MR"
           lgStrSQL =     "SELECT  TOP " & iSelCount  
           lgStrSQL = lgStrSQL & "   dbo.ufn_H_GetCodeName(" & FilterVar("B_BIZ_AREA", "''", "S") & ",y.biz_area_cd,'') BIZ_AREA_nm, "
           lgStrSQL = lgStrSQL & "   y.biz_area_cd, c.emp_no, c.name,a.res_no,c.dept_cd, "

'           lgStrSQL = lgStrSQL & "    bank_accnt,bank,b.bank_nm "

           lgStrSQL = lgStrSQL & " CASE bank_flag "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("1", "''", "S") & "  THEN bank_accnt  "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("2", "''", "S") & " THEN bank_accnt2 "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("3", "''", "S") & " THEN bank_accnt3 "
           lgStrSQL = lgStrSQL & "     END AS bank_accnt, "
           
           lgStrSQL = lgStrSQL & " CASE bank_flag "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("1", "''", "S") & "  THEN bank  "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("2", "''", "S") & " THEN bank2 "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("3", "''", "S") & " THEN bank3 "
           lgStrSQL = lgStrSQL & "     END AS bank, "

           lgStrSQL = lgStrSQL & " CASE bank_flag "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("1", "''", "S") & "  THEN dbo.ufn_H_GetCodeName(" & FilterVar("B_BANK", "''", "S") & ", bank  ,'') "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("2", "''", "S") & " THEN dbo.ufn_H_GetCodeName(" & FilterVar("B_BANK", "''", "S") & ", bank2 ,'') "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("3", "''", "S") & " THEN dbo.ufn_H_GetCodeName(" & FilterVar("B_BANK", "''", "S") & ", bank3 ,'') "
           lgStrSQL = lgStrSQL & "     END AS bank_nm, "

           lgStrSQL = lgStrSQL & " CASE bank_flag "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("1", "''", "S") & "  THEN bankmaster  "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("2", "''", "S") & " THEN bankmaster2 "
           lgStrSQL = lgStrSQL & "     WHEN " & FilterVar("3", "''", "S") & " THEN bankmaster3 "
           lgStrSQL = lgStrSQL & "     END AS bankmaster "
           
           lgStrSQL = lgStrSQL & " From B_COST_CENTER y  ,HDF020T c,b_bank b,haa010t a"
           lgStrSQL = lgStrSQL & " Where " & pCode
           lgStrSQL = lgStrSQL & "   c.dept_cd IN (SELECT x.dept_cd FROM b_acct_dept x, b_company z"
           lgStrSQL = lgStrSQL & "   WHERE x.org_change_id = z.cur_org_change_id"
           lgStrSQL = lgStrSQL & "   AND x.dept_cd = c.dept_cd    AND x.cost_cd = y.cost_cd)"
           lgStrSQL = lgStrSQL & "   AND c.bank=b.bank_cd AND c.emp_no=a.emp_no and"
           lgStrSQL = lgStrSQL & "   c.emp_no in (select  emp_no  from haa010t where retire_dt is null)"

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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
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
                .ggoSpread.SSShowData "<%=lgstrData%>"                               '¢Ð : Display data                
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .DBQueryOk        
           End with
          End If   
    End Select    
       
</Script>	
