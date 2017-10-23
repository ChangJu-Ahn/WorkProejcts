<%@ Language=VBScript%>
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
    
    'On Error Resume Next
    Err.Clear
   
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 
	
	Dim C_SEQ_NO
	Dim C_W18
	Dim C_W19
	Dim C_W20


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
    C_W18		= 2
    C_W19		= 3
    C_W20		= 4
End Sub


'========================================================================================
Sub SubBizQuery()
	Dim iKey1, iKey2, iKey3, blnData1, blnData2
    Dim iDx, iStrData
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	blnData1 = False : blnData2 = False
    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    
    Call SubMakeSQLStatements("D2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements
	If   FncOpenRs("D",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
		Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
		Exit Sub
    End IF
    
    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

        iDx = 1
        'PrintRs lgObjRs
        
        Do While Not lgObjRs.EOF
        
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W1"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W2"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W3"))	
			iStrData = iStrData & Chr(11) 'W22
			iStrData = iStrData & Chr(11) 'W23
			iStrData = iStrData & Chr(11) 'W25
			iStrData = iStrData & Chr(11) 'W26
			iStrData = iStrData & Chr(11) 'W28
			iStrData = iStrData & Chr(11) 'W29
			iStrData = iStrData & Chr(11) 'W30
			iStrData = iStrData & Chr(11) 'W31
			iStrData = iStrData & Chr(11) & iDx + 1
			iStrData = iStrData & Chr(11) & Chr(12)
			iDx = iDx + 1
		    lgObjRs.MoveNext
              
        Loop 
        
        lgObjRs.Close
		blnData1 = True
    End If
        
	Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = True Then

		' 첫번째 쿼리 출력 
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write "" & vbCrLf
		Response.Write "" & vbCrLf
	
		Response.Write " With parent" & vbCr   
		Response.Write "	.ggoSpread.Source = .frm1.vspdData" & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
		Response.Write "	If .frm1.vspdData.MaxRows > 0 Then Call .SetSpreadLock()" & vbCrLf
		Response.Write "	If .frm1.vspdData.MaxRows > 0 Then Call .InsertTotalLine()" & vbCrLf
		Response.Write "	If .frm1.vspdData.MaxRows > 0 Then Call .ChangeRowFlg()" & vbCrLf
		Response.Write " End With"  & vbCrLf
						
		' 두번째 쿼리 출력 
		Response.Write " With parent.frm1" & vbCr   

		Response.Write "	.txtW2.value = " & lgObjRs("W2") & vbCrLf
		Response.Write "	.txtW3.value = " & lgObjRs("W3") & vbCrLf
		Response.Write "	.txtW6.value = " & lgObjRs("W6") & vbCrLf
		Response.Write "	.txtW11.value = " & lgObjRs("W11") & vbCrLf
		Response.Write "	.txtW12.value = " & lgObjRs("W12") & vbCrLf
		Response.Write "	.txtW13.value = " & lgObjRs("W13") & vbCrLf
		Response.Write "	Call parent.SetAllW30_W31"  & vbCrLf
		Response.Write "	Call parent.SetAllRecalc"  & vbCrLf
			
		Response.Write " End With"  & vbCrLf
		Response.Write " </Script>                                          " & vbCr
		blnData2 = True
	End If
	
	If blnData1 = False And blnData2 = False Then
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
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
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W1, A.W2, A.W3 "
            lgStrSQL = lgStrSQL & " FROM TB_LOAN_CALC_SUM A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	ORDER BY A.SEQ_NO ASC" & vbCrLf
	
      Case "R2"
			lgStrSQL =			  " EXEC dbo.usp_TB_26A_GetRef " & pCode1 & "," & pCode2 & "," & pCode3 & " "  
	  Case "D2"
			lgStrSQL =			  " EXEC dbo.usp_TB_26AD_Del " & pCode1 & "," & pCode2 & "," & pCode3 & " "  
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