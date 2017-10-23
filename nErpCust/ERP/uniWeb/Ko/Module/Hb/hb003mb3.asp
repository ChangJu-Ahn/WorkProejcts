<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    
    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    lgSvrDateTime = GetSvrDateTime

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveSingle()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

   
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizSaveSingle()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status
    

    lgStrSQL = ""
    lgStrSQL = lgStrSQL & " DECLARE @DAILY_DED_AMT NUMERIC(18,4) ," & vbCrLf
    lgStrSQL = lgStrSQL & "			@DAILY_TAX_RATE  NUMERIC(18,4) ,"  & vbCrLf
    lgStrSQL = lgStrSQL & "			@DAILY_TAX_DED_RATE NUMERIC(18,4) "  & vbCrLf

    
    lgStrSQL = lgStrSQL & " SELECT  @DAILY_DED_AMT = DAILY_DED_AMT,"  & vbCrLf	' -- 일일근로소득공제금액 
    lgStrSQL = lgStrSQL & "			@DAILY_TAX_RATE = DAILY_TAX_RATE,"  & vbCrLf	' -- 산출세액 
    lgStrSQL = lgStrSQL & "			@DAILY_TAX_DED_RATE = DAILY_TAX_DED_RATE"  & vbCrLf	' -- 세액공제율 
    lgStrSQL = lgStrSQL & " FROM  HFA021T"  & vbCrLf
    lgStrSQL = lgStrSQL & " WHERE COMP_CD = '1' "  & vbCrLf
    lgStrSQL = lgStrSQL & "		AND USE_DT = ("  & vbCrLf
    lgStrSQL = lgStrSQL & "			SELECT MAX(USE_DT)"  & vbCrLf
    lgStrSQL = lgStrSQL & "			FROM HFA021T "  & vbCrLf
    lgStrSQL = lgStrSQL & "			WHERE USE_DT <= " & FilterVar(UCase(lgKeyStream(0)), "''", "S") & vbCrLf	' -- 기준년월이 최근인것.
    lgStrSQL = lgStrSQL & "		)"  & vbCrLf
    lgStrSQL = lgStrSQL & "  "  & vbCrLf
    lgStrSQL = lgStrSQL & "  "  & vbCrLf
    lgStrSQL = lgStrSQL & " UPDATE  HDF071T "  & vbCrLf
    lgStrSQL = lgStrSQL & " SET "  & vbCrLf
    
    ' -- 일당이 공제금액보다 클 경우 차액을 산출세액으로 정리한다.
    lgStrSQL = lgStrSQL & " INCOME_TAX = CASE WHEN B.DAY_MONEY > @DAILY_DED_AMT THEN FLOOR((B.DAY_MONEY - @DAILY_DED_AMT) * @DAILY_TAX_RATE *1.0 /100 * (100 - @DAILY_TAX_DED_RATE) /100 * A.DUTY_DAY /10) *10 ELSE 0 END ," & vbCrLf
    lgStrSQL = lgStrSQL & " RES_TAX    = CASE WHEN B.DAY_MONEY > @DAILY_DED_AMT THEN FLOOR((B.DAY_MONEY - @DAILY_DED_AMT) * @DAILY_TAX_RATE *1.0 /100 * (100 - @DAILY_TAX_DED_RATE) /100 * A.DUTY_DAY /100) *10 ELSE 0 END," & vbCrLf
	
	lgStrSQL = lgStrSQL & " PROV_TOT_AMT = B.DAY_MONEY * A.DUTY_DAY , "  & vbCrLf
	lgStrSQL = lgStrSQL & " REAL_PROV_AMT = PROV_TOT_AMT - SUB_TOT_AMT - INCOME_TAX - RES_TAX, "  & vbCrLf
    lgStrSQL = lgStrSQL & " Updt_emp_no = " & FilterVar(gUsrId, "''", "S")   & "," & vbCrLf
    lgStrSQL = lgStrSQL & " Updt_dt = " & FilterVar(lgSvrDateTime, "''", "S")   & vbCrLf
	
	lgStrSQL = lgStrSQL & " FROM  HDF071T A"  & vbCrLf
	lgStrSQL = lgStrSQL & "		INNER JOIN HAA011T B ON A.EMP_NO = B.EMP_NO"  & vbCrLf
	
    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & "         A.emp_no IN (" & Request("txtSpread") & " ) " & vbCrLf
    lgStrSQL = lgStrSQL & "		AND A.PAY_YYMM = " & FilterVar(UCase(lgKeyStream(0)), "''", "S")     & " "

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

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
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
    
       
</Script>
