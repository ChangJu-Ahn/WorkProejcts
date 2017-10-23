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
    lgStrPrevKey      = UNICInt(Trim(Request("lgStrPrevKey")),0)                    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR","",C_LIKE)

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs, C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx = 1

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & UNIMonthClientFormat(Left(lgObjRs("pay_yymm"),4) & gAPDateSeperator & Right(lgObjRs("pay_yymm"),2) & gAPDateSeperator & "01")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("prov_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("prov_type_nm"))
            lgstrData = lgstrData & Chr(11) & UniDateClientFormat(lgObjRs("prov_dt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("biz_area_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("trans_flag"))
            lgstrData = lgstrData & Chr(11) & UniDateClientFormat(lgObjRs("trans_dt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gl_no"))
    
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
' Desc : Make SQL statements'("MR",iKey1,"X",C_EQ), ("MR",iKey1,"X",C_LIKE)
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pComp)
     Dim iSelCount

     iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

     Select Case  pDataType 
        Case "MR"
           lgStrSQL =     "SELECT  TOP " & iSelCount  
           lgStrSQL = lgStrSQL & "   pay_yymm, prov_type, dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", prov_type) prov_type_nm, prov_dt, "
           lgStrSQL = lgStrSQL & "   dbo.ufn_H_GetCodeName(" & FilterVar("B_BIZ_AREA", "''", "S") & ", biz_area_cd, '') biz_area_nm, "
           lgStrSQL = lgStrSQL & "   CASE trans_flag WHEN " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("생성", "''", "S") & "  ELSE " & FilterVar("취소", "''", "S") & "  END AS trans_flag,"
           lgStrSQL = lgStrSQL & "   trans_dt, gl_no"
           lgStrSQL = lgStrSQL & " From HDF100T"
           lgStrSQL = lgStrSQL & " Where prov_dt between " & FilterVar(lgKeyStream(0),"", "S")
           lgStrSQL = lgStrSQL &        "    and " & FilterVar(lgKeyStream(1),"", "S")
           lgStrSQL = lgStrSQL & "   And prov_type LIKE " & FilterVar(lgKeyStream(2),"'%'", "S")
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
