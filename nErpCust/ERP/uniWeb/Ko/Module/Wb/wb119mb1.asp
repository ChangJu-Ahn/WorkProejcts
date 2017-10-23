<%@ Transaction=required Language=VBScript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    
    On Error Resume Next
    Err.Clear
 
	Dim lgFISC_YEAR1, lgREP_TYPE1
	Dim lgFISC_YEAR2, lgREP_TYPE2

	lgErrorStatus    = "NO"
	                                       
    lgFISC_YEAR1		= Request("txtFISC_YEAR1")
    lgREP_TYPE1			= Request("cboREP_TYPE1")
	lgFISC_YEAR2		= Request("txtFISC_YEAR2")
    lgREP_TYPE2			= Request("cboREP_TYPE2")

	If lgFISC_YEAR1 = "" Or lgFISC_YEAR2 = "" Or _
		lgREP_TYPE1 = "" Or lgREP_TYPE1 = "" Then 

		Response.End 
	Else
		lgOpModeCRUD = UID_M0002
		
		Call SubBizSave()
	End If

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
	Dim arrRowVal, iKey1, iKey2, iKey3
    Dim arrColVal, lgLngMaxRow
    Dim iDx , iType

	Call SubOpenDB(lgObjConn) 
	
	Call CheckVersion(lgFISC_YEAR1, lgREP_TYPE1)	' 2005-03-11 버전관리기능 추가 
	
    'On Error Resume Next
    Err.Clear 

    iKey1 = FilterVar(Request("txtCO_CD2"),"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(lgFISC_YEAR2,"''", "S")	' 사업연도 
	iKey3 = FilterVar(lgREP_TYPE2,"''", "S")	' 사업연도 

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
        Call Displaymsgbox("WC0040", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
    
		lgStrSQL = "EXEC dbo.usp_WB119MB1_CopyVersion" & Request("rdoType")
		
		lgStrSQL = lgStrSQL & " '" & Request("txtCO_CD2")	& "'"  & vbCrLf 
		lgStrSQL = lgStrSQL & ",'" & lgFISC_YEAR2	& "'"  & vbCrLf 
		lgStrSQL = lgStrSQL & ",'" & lgREP_TYPE2	& "'"  & vbCrLf 

		lgStrSQL = lgStrSQL & ",'" & wgCO_CD		& "'"  & vbCrLf 
		lgStrSQL = lgStrSQL & ",'" & lgFISC_YEAR1	& "'"  & vbCrLf 
		lgStrSQL = lgStrSQL & ",'" & lgREP_TYPE1	& "'"  & vbCrLf 

		lgStrSQL = lgStrSQL & ",'" & gUsrID		& "'"  & vbCrLf 
	
		PrintLog "SubBizSave = " & lgStrSQL
	
		lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
		
	End If
	
	Call SubCloseDB(lgObjConn)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " TOP 1 1 "
            lgStrSQL = lgStrSQL & " FROM TB_COMPANY_HISTORY A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
    End Select
	PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
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
    lgErrorStatus     = "YES"
End Sub

'========================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    'On Error Resume Next
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

       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>
<%
'   **************************************************************
'	1.4 Transaction 처러 이벤트 
'   **************************************************************

Sub	onTransactionCommit()
	' 트랜잭션 완료후 이벤트 처리 
End Sub

Sub onTransactionAbort()
	' 트랜잭선 실패(에러)후 이벤트 처리 
'PrintForm
'	' 에러 출력 
	'Call SaveErrorLog(Err)	' 에러로그를 남긴 
	
End Sub
%>
