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
    
    'On Error Resume Next
    Err.Clear
   
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	Dim C_W22
	Dim C_W23
	Dim C_W23_NM
	Dim C_W24
	Dim C_W25
	Dim C_W26
	Dim C_W28
	Dim C_W29
	Dim C_W31
	Dim C_W32


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
	C_W22 = 1
	C_W23 = 2
	C_W23_NM = 3
	C_W24 = 4
	C_W25 = 5
	C_W26 = 6
	C_W28 = 7
	C_W29 = 8
	C_W31 = 9
	C_W32 = 10
End Sub


'========================================================================================
Sub SubBizQuery()
	Dim iKey1, iKey2, iKey3
    Dim iDx, iStrData, blnNoData1, blnNoData2
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	blnNoData1 = True : blnNoData2 = True
	
    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
	
	Call SubMakeSQLStatements("D",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements
	If   FncOpenRs("D",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
		Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
		Exit Sub
    End IF
        
    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
        blnNoData1 = False
    Else
        iDx = 1
        'PrintRs lgObjRs
        
        Do While Not lgObjRs.EOF
        
			iStrData = iStrData & Chr(11) & ConvSPChars(iDx)
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W22"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W23"))	
			iStrData = iStrData & Chr(11)	'W23_BT
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W23_NM"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W24"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W25"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W26"))	
			iStrData = iStrData & Chr(11)	'W27
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W28"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W29"))	
			iStrData = iStrData & Chr(11)	'W30
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W31"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W32"))	
			iStrData = iStrData & Chr(11) 'DESC
			iStrData = iStrData & Chr(11) & iDx + 1
			iStrData = iStrData & Chr(11) & Chr(12)
			iDx = iDx + 1
		    lgObjRs.MoveNext
              
        Loop 
        ' 합계 
		iStrData = iStrData & Chr(11) & "999999"
		iStrData = iStrData & Chr(11)
		iStrData = iStrData & Chr(11)
		iStrData = iStrData & Chr(11)	'W23_BT
		iStrData = iStrData & Chr(11)
		iStrData = iStrData & Chr(11)
		iStrData = iStrData & Chr(11)
		iStrData = iStrData & Chr(11) & "0"
		iStrData = iStrData & Chr(11) & "0"
		iStrData = iStrData & Chr(11) & "0"
		iStrData = iStrData & Chr(11) & "0"
		iStrData = iStrData & Chr(11) & "0"
		iStrData = iStrData & Chr(11) & "0"
		iStrData = iStrData & Chr(11) & "0"
		iStrData = iStrData & Chr(11) 'DESC
		iStrData = iStrData & Chr(11) & iDx + 1
		iStrData = iStrData & Chr(11) & Chr(12)
        
        lgObjRs.Close
	End If
	
	Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
		 blnNoData2 = False  
	Else
		' 첫번째 쿼리 출력 
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write "" & vbCrLf
		Response.Write "" & vbCrLf
	
		Response.Write " With parent" & vbCr   
		Response.Write "	.GetRefOK """ & iStrData       & """" & vbCr
		Response.Write "	If .frm1.vspdData2.MaxRows > 0 Then Call .SetSpreadLock()" & vbCrLf
		Response.Write "	If .frm1.vspdData2.MaxRows > 0 Then Call .QueryTotalLine()" & vbCrLf
		Response.Write "	If .frm1.vspdData2.MaxRows > 0 Then Call .ChangeRowFlg()" & vbCrLf
		Response.Write " End With"  & vbCrLf
						
		' 두번째 쿼리 출력 
		Response.Write " With parent.frm1" & vbCr   
 
		Response.Write "	.txtW1_BEFORE.value = " & lgObjRs("W2") & vbCrLf
		Response.Write "	.txtW4.value = " & lgObjRs("W4") & vbCrLf
		Response.Write "	.txtW4_BEFORE.value = " & lgObjRs("W4") & vbCrLf
		Response.Write "	.txtW6.value = " & lgObjRs("W6") & vbCrLf
		Response.Write "	.txtW8.value = " & lgObjRs("W8") & vbCrLf
		Response.Write "	.txtW10_BEFORE.value = " & lgObjRs("W10") & vbCrLf
		Response.Write "	.txtW14.value = " & lgObjRs("W14") & vbCrLf
		Response.Write "	Call parent.SetAllTxtRecalc " & vbCrLf
			
		Response.Write " End With"  & vbCrLf
		Response.Write " </Script>                                          " & vbCr
	End If

	'PrintLog "lgObjRs(W2) = " & lgObjRs("W2") & "::"
	
	If blnNoData1 = False And blnNoData2 = False Then
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
	  Case "D"
			lgStrSQL=" EXEC dbo.usp_tb_34D2_del " & pCode1 & "," & pCode2 & "," & pCode3 & " "  
      Case "R"
			lgStrSQL =			  " SELECT MAX(A.W2) AS W22, A.W1 AS W23, A.W1_NM AS W23_NM, MAX(A.W3) AS W24, MAX(A.W11) AS W25" & vbCrLf
			lgStrSQL = lgStrSQL & "	, SUM(A.W5) AS W26, SUM( CASE WHEN A.W7 = '1' THEN A.W6 ELSE 0 END) AS W28, SUM( CASE WHEN A.W7 = '2' THEN A.W6 ELSE 0 END) AS W29" & vbCrLf
			lgStrSQL = lgStrSQL & "	, SUM( CASE WHEN A.W9 = '1' THEN A.W8 ELSE 0 END) AS W31, SUM( CASE WHEN A.W9 = '2' THEN A.W8 ELSE 0 END) AS W32 " & vbCrLf
			lgStrSQL = lgStrSQL & "	FROM TB_BED_DEBT A (NOLOCK) " & vbCrLf
			lgStrSQL = lgStrSQL & "	WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	AND A.SEQ_NO <> '999999' " & vbCrLf
			lgStrSQL = lgStrSQL & "	AND ISNULL(A.W1, '') <> '' " & vbCrLf
			lgStrSQL = lgStrSQL & "	GROUP BY A.W1, A.W1_NM " & vbCrLf

      Case "R2"
			lgStrSQL =			  " EXEC dbo.usp_TB_34_GetRef " & pCode1 & "," & pCode2 & "," & pCode3 & " "  
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