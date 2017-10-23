<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<%Option Explicit%> 
<% session.CodePage=949 %>
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

	Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 
	Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.
			
	Dim C_SEQ_NO1
	Dim C_W7
	Dim C_W8
	Dim C_W8_NM
	Dim C_W9
	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W13

	Dim C_SEQ_NO2
	Dim C_W14
	Dim C_W15
	Dim C_W16
	Dim C_W17
	Dim C_W18
	Dim C_W19
	Dim C_W20
	Dim C_W21
	Dim C_HEAD_SEQ_NO1

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

	C_SEQ_NO1	= 1	' -- 1번 그리드 
    C_W7		= 2
    C_W8		= 3
    C_W8_NM		= 4
    C_W9		= 5
    C_W10		= 6
    C_W11		= 7
    C_W12		= 8
    C_W13		= 9	
 
 	C_SEQ_NO2	= 1  ' -- 2번 그리드 
    C_W14		= 2 
    C_W15		= 3
    C_W16		= 4
    C_W17		= 5
    C_W18		= 6
    C_W19		= 7
    C_W20		= 8
    C_W21		= 9
    C_HEAD_SEQ_NO1 = 10
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
	   ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
	    lgstrData = "" : iLngRow = 1 : iRow = TYPE_2
		iMaxRows = lgObjRs.RecordCount
		
		iRow = TYPE_1  
		PrintLog "iMaxRows = " & iMaxRows
		sData = "	.lgCurrGrid = " & TYPE_1 & vbCrLf
		sData = sData & "	Call .FncInsertRow(1)" & vbCrLf
		sData = sData & "	If " & 	iMaxRows & " > 1 Then " & vbCrLf
		sData = sData & "		Call .FncInsertRow(" & 	iMaxRows-1 & ")" & vbCrLf
		sData = sData & "	End If" & vbCrLf
		sData = sData & "	.lgCurrGrid = " & TYPE_2 & vbCrLf
		sData = sData & "	Call .FncInsertRow(1)" & vbCrLf
		sData = sData & "	If " & 	iMaxRows & " > 1 Then " & vbCrLf
		sData = sData & "		Call .FncInsertRow(" & 	iMaxRows-1 & ")" & vbCrLf
		sData = sData & "	End If" & vbCrLf
		sData = sData & "	.lgvspdData(" & TYPE_1 & ").Redraw = False" & vbCrLf
		sData = sData & "	.lgvspdData(" & TYPE_2 & ").Redraw = False" & vbCrLf
					
		Do While Not lgObjRs.EOF
			sData = sData & "	Call .PutGrid(" & TYPE_1 & ", " & C_W7 & ", " & iLngRow & ", """ & lgObjRs("COMPANY_NM") & """)" & vbCrLf
			sData = sData & "	Call .PutGridText(" & TYPE_1 & ", " & C_W8 & ", " & iLngRow & ", """ & lgObjRs("COMPANY_TYPE") & """)" & vbCrLf
			sData = sData & "	Call .PutGridText(" & TYPE_1 & ", " & C_W8_NM & ", " & iLngRow & ", """ & lgObjRs("COMPANY_TYPE_NM") & """)" & vbCrLf
			sData = sData & "	Call .PutGrid(" & TYPE_1 & ", " & C_W9 & ", " & iLngRow & ", """ & lgObjRs("OWN_RGST_NO") & """)" & vbCrLf
			sData = sData & "	Call .PutGrid(" & TYPE_1 & ", " & C_W10 & ", " & iLngRow & ", """ & lgObjRs("CO_ADDR") & """)" & vbCrLf
			sData = sData & "	Call .PutGrid(" & TYPE_1 & ", " & C_W11 & ", " & iLngRow & ", """ & lgObjRs("REPRE_NM") & """)" & vbCrLf
			sData = sData & "	Call .PutGrid(" & TYPE_1 & ", " & C_W12 & ", " & iLngRow & ", """ & lgObjRs("STOCK_CNT") & """)" & vbCrLf
			sData = sData & "	Call .PutGrid(" & TYPE_1 & ", " & C_W13 & ", " & iLngRow & ", """ & lgObjRs("STOCK_RATE") & """)" & vbCrLf
			
			sData = sData & "	Call .PutGrid(" & TYPE_2 & ", " & C_W14 & ", " & iLngRow & ", """ & lgObjRs("COMPANY_NM") & """)" & vbCrLf
			sData = sData & "	Call .PutGrid(" & TYPE_2 & ", " & C_W15 & ", " & iLngRow & ", """ & lgObjRs("DOC_AMT") & """)" & vbCrLf
			sData = sData & "	Call .PutGridText(" & TYPE_2 & ", " & C_W16 & ", " & iLngRow & ", """ & lgObjRs("w16") & """)" & vbCrLf
			sData = sData & "	Call .PutGrid(" & TYPE_2 & ", " & C_W17 & ", " & iLngRow & ", """ & lgObjRs("w17") & """)" & vbCrLf
			sData = sData & "	Call .PutGrid(" & TYPE_2 & ", " & C_W19 & ", " & iLngRow & ", """ & lgObjRs("w19") & """)" & vbCrLf
			sData = sData & "	Call .PutGrid(" & TYPE_2 & ", " & C_HEAD_SEQ_NO1 & ", " & iLngRow & ", """ & iLngRow & """)" & vbCrLf
			iLngRow = iLngRow + 1
			lgObjRs.MoveNext
		Loop 

		lgObjRs.Close
		Set lgObjRs = Nothing
		
		' -- 제26호 지급이자/대차대조표상의 자산총액을 구해온다 
		Call SubMakeSQLStatements("D",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then
			' -- 전역변수에 저장 
			sData = sData & "	.lgTB_26_AMT = " & lgObjRs("TB_26_AMT") & vbCrLf
			sData = sData & "	.lgTB_3_AMT = " & lgObjRs("TB_3_AMT") & vbCrLf
			
		End If
		
		sData = sData & "	Call .ReClacGrid2 " & vbCrLf
		sData = sData & "	.lgvspdData(" & TYPE_1 & ").Redraw = True" & vbCrLf
		sData = sData & "	.lgvspdData(" & TYPE_2 & ").Redraw = True" & vbCrLf
		'sData = sData & "	Call .ReClacGrid2()" & vbCrLf

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
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
			lgStrSQL = lgStrSQL & " DECLARE @W16	NUMERIC(15) " & vbCrLf
			lgStrSQL = lgStrSQL & "		,	@TB_26_AMT	NUMERIC(15) " & vbCrLf
			lgStrSQL = lgStrSQL & "		,	@TB_3_2_AMT	NUMERIC(15) " & vbCrLf
			
			lgStrSQL = lgStrSQL & " SELECT @TB_3_2_AMT = dbo.ufn_TB_16_2_GetRef2(" & pCode1 & ", " & pCode2 & ", " & pCode3 & ") " & vbCrLf	' -- 대차대조표상의 자산합계 
			lgStrSQL = lgStrSQL & " SELECT @TB_26_AMT = dbo.ufn_TB_16_2_GetRef3(" & pCode1 & ", " & pCode2 & ", " & pCode3 & ") " & vbCrLf	' -- 제26호갑의 지급이자 

			lgStrSQL = lgStrSQL & " SELECT dbo.ufn_TB_16_2_GetRef(" & pCode1 & ", " & pCode2 & ", " & pCode3 & ", A.STOCK_RATE) W16 " & vbCrLf ' -- 입금불산입율 표시분	' -- 익금불산입 
            lgStrSQL = lgStrSQL & " , A.COMPANY_NM, A.COMPANY_TYPE, dbo.ufn_GetCodeName('W1004', A.COMPANY_TYPE) COMPANY_TYPE_NM, A.OWN_RGST_NO "& vbCrLf
            lgStrSQL = lgStrSQL & " , A.CO_ADDR, A.REPRE_NM, A.STOCK_CNT, A.STOCK_RATE, A.DOC_AMT " & vbCrLf
            lgStrSQL = lgStrSQL & " , A.DOC_AMT * (dbo.ufn_TB_16_2_GetRef(" & pCode1 & ", " & pCode2 & ", " & pCode3 & ", A.STOCK_RATE) / 100) W17 "& vbCrLf
            lgStrSQL = lgStrSQL & " , @TB_26_AMT * ( (A.ACQUIRE_AMT * A.HOLDING_TERM * (dbo.ufn_TB_16_2_GetRef(" & pCode1 & ", " & pCode2 & ", " & pCode3 & ", A.STOCK_RATE) /100)) / (@TB_3_2_AMT * 365 )	) W19 " & vbCrLf ' -- 지급이자*(취득가액*보유기간 /자산합계)
            lgStrSQL = lgStrSQL & " FROM TB_DIVIDEND A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	ORDER BY A.DOC_DATE DESC " & vbCrLf
			
      Case "D"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "		dbo.ufn_TB_16_2_GetRef3(" & pCode1 & ", " & pCode2 & ", " & pCode3 & ") TB_26_AMT" & vbCrLf
            lgStrSQL = lgStrSQL & "	,	dbo.ufn_TB_16_2_GetRef2(" & pCode1 & ", " & pCode2 & ", " & pCode3 & ") TB_3_AMT" & vbCrLf

	
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