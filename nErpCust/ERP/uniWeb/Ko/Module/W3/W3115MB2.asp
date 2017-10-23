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
    Call CheckVersion(sFISC_YEAR, sREP_TYPE) 	
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
	Dim iKey1, iKey2, iKey3
    Dim iDx, iStrData, iStrData2, iStrData3
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
        iDx = 1
        'PrintRs lgObjRs
        
        Do While Not lgObjRs.EOF
        
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W1"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W2"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W3"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W4"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W5"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("R1"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("R2"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("R3"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("R4"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("R5"))	
			iStrData = iStrData & Chr(11) & iDx + 1
			iStrData = iStrData & Chr(11) & Chr(12)
			iDx = iDx + 1
		    lgObjRs.MoveNext
              
        Loop 
        
        lgObjRs.Close
	    
		Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
		     lgStrPrevKey = ""
		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		    Call SetErrorStatus()
		    
		Else

	        iDx = 1
	        'PrintRs lgObjRs
	        
	        Do While Not lgObjRs.EOF
	        
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("CHILD_SEQ_NO"))			
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W6"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W7"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W8"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W9"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W10"))	
				iStrData2 = iStrData2 & Chr(11) & iDx + 1
				iStrData2 = iStrData2 & Chr(11) & Chr(12)
				iDx = iDx + 1
			    lgObjRs.MoveNext
	              
	        Loop 
	        
	        lgObjRs.Close

		End If

    End If

    Call SubMakeSQLStatements("R3",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
        iDx = 1
        'PrintRs lgObjRs
        
        Do While Not lgObjRs.EOF
        
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W_TYPE"))
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W11"))			
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W12"))	
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W13"))	
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W14"))	
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W15"))			
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W16"))	
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W17"))	
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W18"))	
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W19"))			
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W20"))	
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W21"))	
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W22"))	
			iStrData3 = iStrData3 & Chr(11) & ConvSPChars(lgObjRs("W23"))	
			iStrData3 = iStrData3 & Chr(11) 	
			iStrData3 = iStrData3 & Chr(11) & iDx + 1
			iStrData3 = iStrData3 & Chr(11) & Chr(12)
			iDx = iDx + 1
		    lgObjRs.MoveNext
              
        Loop 
        
        lgObjRs.Close
	    
    End If

	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData2       & """" & vbCr	

    Response.Write "	.ggoSpread.Source = .frm1.vspdData3              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData3       & """" & vbCr
    Response.Write "	Call .SetInitGrid3              " & vbCr
    
    If lgErrorStatus = "NO" Then
		Response.Write "	.GetRefOk                                      " & vbCr
	End If
    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>                                          " & vbCr
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
	  Case "D"		'add 20060126 byHJO
			lgStrSQL=" EXEC dbo.usp_tb_19A_del " & pCode1 & "," & pCode2 & "," & pCode3 & " "  
      Case "R"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " SEQ_NO, W1, W2, W3, W4, W5, R1, R2, R3, R4, R5 "
            lgStrSQL = lgStrSQL & " FROM DBO.ufn_TB_19A_GetRef( " & pCode1 	 & "," & pCode2 	 & "," & pCode3 	& ") "	 & vbCrLf
	
      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " SEQ_NO, CHILD_SEQ_NO, W6, W7, W8, W9, W10 "
            lgStrSQL = lgStrSQL & " FROM DBO.ufn_TB_19A_GetRef2( " & pCode1 	 & "," & pCode2 	 & "," & pCode3 	& ") "	 & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY SEQ_NO ASC, CHILD_SEQ_NO ASC " & VbCrLf

      Case "R3"
			lgStrSQL =			  " SELECT W_TYPE, SEQ_NO, W11, W12, W13 / 100 AS W13, W14 / 100 AS W14, W15 / 100 AS W15 " & VbCrLf
            lgStrSQL = lgStrSQL & " 	, W16 / 100 AS W16, W17 / 100 AS W17, W18 / 100 AS W18, W19 / 100 AS W19, W20 / 100 AS W20, W21 / 100 AS W21, W22 / 100 AS W22, W23 / 100 AS W23 " & VbCrLf
            lgStrSQL = lgStrSQL & " FROM DBO.ufn_TB_19A_GetRef3_1( " & pCode1 	 & "," & pCode2 	 & "," & pCode3 	& ") "	 & vbCrLf
            lgStrSQL = lgStrSQL & " UNION  " & VbCrLf
            lgStrSQL = lgStrSQL & " SELECT W_TYPE, SEQ_NO, W11, W12, W13, W14, W15 " & VbCrLf
            lgStrSQL = lgStrSQL & " 	, W16, W17, W18, W19, W20, W21, W22, W23 " & VbCrLf
            lgStrSQL = lgStrSQL & " FROM DBO.ufn_TB_19A_GetRef3_2( " & pCode1 	 & "," & pCode2 	 & "," & pCode3 	& ") "	 & vbCrLf
            lgStrSQL = lgStrSQL & " UNION  " & VbCrLf
            lgStrSQL = lgStrSQL & " SELECT 'D' AS W_TYPE, SEQ_NO " & VbCrLf
            lgStrSQL = lgStrSQL & "		, CASE WHEN SEQ_NO = 999999 THEN '계' ELSE W1 END AS W11, W5 AS W12 " & VbCrLf
            lgStrSQL = lgStrSQL & " 	, 0 AS W13, 0 AS W14, 0 AS W15, 0 AS W16, 0 AS W17 " & VbCrLf
            lgStrSQL = lgStrSQL & " 	, 0 AS W18, 0 AS W19, 0 AS W20, 0 AS W21, 0 AS W22, 0 AS W22 " & VbCrLf
            lgStrSQL = lgStrSQL & " FROM DBO.ufn_TB_19A_GetRef( " & pCode1 	 & "," & pCode2 	 & "," & pCode3 	& ") "	 & vbCrLf
'            lgStrSQL = lgStrSQL & " WHERE SEQ_NO <> 999999 " & VbCrLf

            lgStrSQL = lgStrSQL & " ORDER BY W_TYPE DESC, SEQ_NO " & VbCrLf

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