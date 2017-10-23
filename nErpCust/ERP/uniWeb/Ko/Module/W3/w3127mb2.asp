<%@ Language=VBScript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

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
	Const TYPE_3	= 2		
	Const TYPE_4A	= 3
	Const TYPE_4B	= 4
	Const TYPE_5	= 5
	Const TYPE_6	= 6
	
	Dim C_SEQ_NO
	Dim C_W_TYPE
	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7

	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W13
	Dim C_W14

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
	C_SEQ_NO	= 1	' -- 1번 그리드 
    C_W_TYPE	= 2	' 구분 
    C_W1		= 3	' 연월일 
    C_W2		= 4 ' 적요 
    C_W3		= 5	' 차변 
    C_W4		= 6	' 대변 
    C_W5		= 7	' 잔액 
    C_W6		= 8	' 일수 
    C_W7		= 9	' 적수 

	C_W10		= 5	' 자산총계 
	C_W11		= 6	' 부채총계 
	C_W12		= 7	' 자기자본 
	C_W13		= 8 ' 연일수 
	C_W14		= 9	' 적수 
End Sub


'========================================================================================
Sub SubBizQuery()
	Dim iKey1, iKey2, iKey3
    Dim iDx, arrRs(6)
    Dim iRow, sW2, sW3, sW10, sW11, sW12, sW12_REF
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
		sW2		= lgObjRs("W2") 'UNINumClientFormat(lgObjRs("W2"), ggAmtOfMoney.DecPoint, 0)
		sW3		= lgObjRs("W3") 'UNINumClientFormat(lgObjRs("W3"), ggAmtOfMoney.DecPoint, 0)
		sW10	= lgObjRs("W10") 'UNINumClientFormat(lgObjRs("W10"), ggAmtOfMoney.DecPoint, 0)
		sW11	= lgObjRs("W11") 'UNINumClientFormat(lgObjRs("W11"), ggAmtOfMoney.DecPoint, 0)
		sW12	= lgObjRs("W12") 'UNINumClientFormat(lgObjRs("W12"), ggAmtOfMoney.DecPoint, 0)
		sW12_REF	= lgObjRs("W12_REF") 'UNINumClientFormat(lgObjRs("W12_REF"), ggAmtOfMoney.DecPoint, 0)
    End If
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write "" & vbCrLf
	Response.Write "" & vbCrLf
	
	Response.Write " With parent.lgvspdData(" & TYPE_4A & ")" & vbCr   
'
	Response.Write " parent.lgCurrGrid = " & TYPE_4A & vbCrLf
	Response.Write " If .MaxRows = 0 Then	Call parent.InsertTotalLine(" & TYPE_4A & ")" & vbCrLf
	Response.Write "	.Col = parent.C_W7 : .Row = .MaxRows : .text = """ & sW2 & """" & vbCrLf
	Response.Write " End With"  & vbCrLf
'
	Response.Write " With parent.lgvspdData(" & TYPE_4B & ")" & vbCr   
	Response.Write " parent.lgCurrGrid = " & TYPE_4B & vbCrLf
	Response.Write " If .MaxRows = 0 Then	Call parent.InsertTotalLine(" & TYPE_4B & ")" & vbCrLf
	Response.Write "	.Col = parent.C_W7 : .Row = .MaxRows : .text = """ & sW3 & """" & vbCrLf
	Response.Write " End With"  & vbCrLf
'
	Response.Write " With parent.lgvspdData(" & TYPE_6 & ")" & vbCr   
	Response.Write " parent.lgCurrGrid = " & TYPE_6 & vbCrLf
	'Response.Write " If .MaxRows <= 1 Then	Call parent.fncInsertRow(1)" & vbCrLf  ''?
	Response.Write "	.Row = .MaxRows " & vbCrLf
	Response.Write "	.Col = parent.C_W10 : .text = """ & sW10 & """" & vbCrLf
	Response.Write "	.Col = parent.C_W11 : .text = """ & sW11 & """" & vbCrLf
	Response.Write "	.Col = parent.C_W12 : .text = """ & sW12 & """" & vbCrLf
	Response.Write " End With"  & vbCrLf
	Response.Write " parent.lgW12_REF = " & sW12_REF & vbCrLf	
	Response.Write " Call parent.SetW14()      "& vbCrLf
    Response.Write " </Script>                                          " & vbCr

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"
			lgStrSQL =			  " EXEC dbo.usp_TB_26B_GetRef " & pCode1 & "," & pCode2 & "," & pCode3 & " "  
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