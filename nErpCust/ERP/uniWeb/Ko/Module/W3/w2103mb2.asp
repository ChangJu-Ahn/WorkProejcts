<%@ Language=VBScript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../wcm/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    
    'On Error Resume Next
    Err.Clear
   
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	Dim C_SEQ_NO
	Dim C_DOC_DATE
	Dim C_DOC_AMT
	Dim C_DEBIT_CREDIT
	Dim C_DEBIT_CREDIT_NM
	Dim C_SUMMARY_DESC
	Dim C_COMPANY_NM
	Dim C_STOCK_RATE
	Dim C_ACQUIRE_AMT
	Dim C_COMPANY_TYPE
	Dim C_COMPANY_TYPE_NM
	Dim C_HOLDING_TERM
	Dim C_JUKSU
	Dim C_OWN_RGST_NO
	Dim C_CO_ADDR
	Dim C_REPRE_NM
	Dim C_STOCK_CNT

	lgErrorStatus    = "NO"
    lgOpModeCRUD     = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    'lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    'lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    'lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

    Call SubOpenDB(lgObjConn) 
    	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
    End Select

    Call SubCloseDB(lgObjConn)
    
    Response.End 
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
	C_SEQ_NO			= 1
    C_DOC_DATE			= 2
    C_DOC_AMT			= 3
    C_DEBIT_CREDIT		= 4
    C_DEBIT_CREDIT_NM	= 5
    C_SUMMARY_DESC		= 6
    C_COMPANY_NM		= 7
    C_STOCK_RATE		= 8
    C_ACQUIRE_AMT		= 9
    C_COMPANY_TYPE		= 10
    C_COMPANY_TYPE_NM	= 11
    C_HOLDING_TERM		= 12
    C_JUKSU				= 13
    C_OWN_RGST_NO		= 14
    C_CO_ADDR			= 15
    C_REPRE_NM			= 16
    C_STOCK_CNT			= 17
End Sub


'========================================================================================
Sub SubBizQuery()
	Dim iKey1, iKey2, iKey3, iMaxRows, iLngRow
    Dim iDx, sData
    Dim iRow, sW2, sW3, sW10, sW11, sW12, sW12_REF
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

	' 2번 그리드 
	Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	gCursorLocation = adUseClient
	If   FncOpenRs("P",lgObjConn,lgObjRs,lgStrSQL, adOpenKeyset, adLockReadOnly) = False Then
  
	     lgStrPrevKey = ""
	    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
		    
	Else
	    lgstrData = "" : iLngRow = 1 
		iMaxRows = lgObjRs.RecordCount
		
		PrintLog "iMaxRows = " & iMaxRows

		sData = sData & "	Call parent.FncInsertRow(" & 	iMaxRows & ")" & vbCrLf

		sData = sData & "	.Redraw = False" & vbCrLf
					
		Do While Not lgObjRs.EOF
			sData = sData & "	.Row = " & iLngRow & vbCrLf
			sData = sData & "	.Col = " & C_DOC_DATE & " : .text =""" & lgObjRs("DOC_DT") & """ " & vbCrLf
			sData = sData & "	.Col = " & C_DOC_AMT & " : .value =""" & lgObjRs("DOC_AMT") & """ " & vbCrLf
			sData = sData & "	.Col = " & C_SUMMARY_DESC & " : .text =""" & lgObjRs("DOC_DESC") & """ " & vbCrLf
			
			iLngRow = iLngRow + 1
			lgObjRs.MoveNext
		Loop 
		sData = sData & "	.Redraw = true" & vbCrLf
		
		lgObjRs.Close
		Set lgObjRs = Nothing
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent.frm1.vspdData                               " & vbCr
		Response.Write sData & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr		
	End If
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)

    Select Case pMode 
      Case "R"
			lgStrSQL = lgStrSQL & " SELECT DOC_DT, DOC_DESC " & vbCrLf
			lgStrSQL = lgStrSQL & "		,	(CASE CREDIT_DEBIT  " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHEN 'DR' THEN DOC_AMT * -1  " & vbCrLf
			lgStrSQL = lgStrSQL & "			WHEN 'CR' THEN DOC_AMT " & vbCrLf
			lgStrSQL = lgStrSQL & "			END) DOC_AMT " & vbCrLf
			lgStrSQL = lgStrSQL & " FROM TB_WORK_3 " & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE ACCT_CD IN ( " & vbCrLf
			lgStrSQL = lgStrSQL & "		SELECT ACCT_CD  "
            lgStrSQL = lgStrSQL & "		FROM TB_ACCT_MATCH " & vbCrLf
            lgStrSQL = lgStrSQL & "		WHERE	CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND MATCH_CD = '05'" 	 & vbCrLf
			lgStrSQL = lgStrSQL & " ) " & vbCrLf
            lgStrSQL = lgStrSQL & "		AND	CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "			AND REP_TYPE = " & pCode3 	 & vbCrLf
    End Select

	PrintLog "SubMakeSQLStatements = " & lgStrSQL
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