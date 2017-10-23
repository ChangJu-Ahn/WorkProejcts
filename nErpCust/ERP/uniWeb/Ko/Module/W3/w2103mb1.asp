<%@ LANGUAGE=VBSCript%>
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
	Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 
	Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.

	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey, lgCurrGrid
	
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

	Dim C_MINOR_NM
	Dim C_MINOR_CD
	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W3_1
	Dim C_W3_2
	Dim C_W3_3
	Dim C_W3_4
	Dim C_W3_5
	Dim C_W4

	lgErrorStatus		= "NO"
    lgOpModeCRUD		= Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR			= Request("txtFISC_YEAR")
    sREP_TYPE			= Request("cboREP_TYPE")
	lgCurrGrid			= Request("txtCurrGrid")
	
    lgPrevNext			= Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    lgLngMaxRow			= Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey		= UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

    Call SubOpenDB(lgObjConn) 
    
    Call CheckVersion(sFISC_YEAR, sREP_TYPE)	' 2005-03-11 버전관리기능 추가 
    	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)
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

    C_MINOR_NM			= 2
    C_MINOR_CD			= 3
    C_W1				= 4
    C_W2				= 5
    C_W3				= 6
    C_W3_1				= 6
    C_W3_2				= 7
    C_W3_3				= 8
    C_W3_4				= 9
    C_W3_5				= 10
    C_W4				= 11
End Sub

'========================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear


	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_DIVIDEND_REF WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_DIVIDEND WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf

PrintLog "SubBizDelete = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)    
End Sub

' 환경변수1 %의 콤보값 인덱스 가져옴 
Function ReadCombo1(pVal)
	On Error Resume Next
	If pVal = Trim("") Then
		ReadCombo1 = 0
	Else
		ReadCombo1 = (UNICDbl(pVal, 0) / 10) + 1	' 공백포함 
	End If
	If Err Then PrintLog "ReadCombo1=" & pVal
End Function

' 환경변수2 초과등 콤보값 인덱스가져옴 
Function ReadCombo2(pVal)
	Select Case pVal
		Case ">"
			ReadCombo2 = "초과"
		Case ">="
			ReadCombo2 = "이상"
		Case "<="
			ReadCombo2 = "이하"
		Case "<"
			ReadCombo2 = "미만"
		Case Else
			ReadCombo2 = " "
	End Select
End Function

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1)
    Dim iDx
    Dim iLoopMax
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

'	If CInt(lgCurrGrid) = TYPE_1 Then
	
		Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
		     lgStrPrevKey = ""
		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		    Call SetErrorStatus()
		    
		Else
		    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				arrRs(0) = arrRs(0) & Chr(11) & UNIDateClientFormat(lgObjRs("DOC_DATE"))	
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("DOC_AMT"))			
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("DEBIT_CREDIT"))	
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("DEBIT_CREDIT_NM"))	
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("SUMMARY_DESC"))	
				'arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("COMPANY_CD"))		
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("COMPANY_NM"))		
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("STOCK_RATE"))		
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("ACQUIRE_AMT"))		
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("COMPANY_TYPE"))	
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("COMPANY_TYPE_NM"))	 
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("HOLDING_TERM"))	
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("JUKSU"))	
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("OWN_RGST_NO"))		
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("CO_ADDR"))		
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("REPRE_NM"))		
				arrRs(0) = arrRs(0) & Chr(11) & ConvSPChars(lgObjRs("STOCK_CNT"))			 
				arrRs(0) = arrRs(0) & Chr(11) & iIntMaxRows + iLngRow + 1
				arrRs(0) = arrRs(0) & Chr(11) & Chr(12)
			    lgObjRs.MoveNext

		        iDx =  iDx + 1
		        'If iDx > C_SHEETMAXROWS_D Then
		        ''   lgStrPrevKey = lgStrPrevKey + 1
		        '   Exit Do
		       ' End If               
		    Loop 
		    
		    lgObjRs.Close
    
		End If
    
'    ElseIf CInt(lgCurrGrid) = TYPE_2 Then
    
		Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then

		    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = ""
			    
		    iDx = 1
			    
		    Do While Not lgObjRs.EOF
				arrRs(1) = arrRs(1) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				arrRs(1) = arrRs(1) & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
				arrRs(1) = arrRs(1) & Chr(11) & ConvSPChars(lgObjRs("MINOR_CD"))
				arrRs(1) = arrRs(1) & Chr(11) & ConvSPChars(lgObjRs("W1"))
				arrRs(1) = arrRs(1) & Chr(11) & ConvSPChars(lgObjRs("W2"))
				arrRs(1) = arrRs(1) & Chr(11) & ConvSPChars(lgObjRs("W3_1"))
				arrRs(1) = arrRs(1) & Chr(11) & ConvSPChars(ReadCombo2(lgObjRs("W3_2")))
				arrRs(1) = arrRs(1) & Chr(11) & "~"
				arrRs(1) = arrRs(1) & Chr(11) & ConvSPChars(lgObjRs("W3_3"))
				arrRs(1) = arrRs(1) & Chr(11) & ConvSPChars(ReadCombo2(lgObjRs("W3_4")))
				arrRs(1) = arrRs(1) & Chr(11) & ConvSPChars(lgObjRs("W4"))
							 
				arrRs(1) = arrRs(1) & Chr(11) & iIntMaxRows + iLngRow + 1
				arrRs(1) = arrRs(1) & Chr(11) & Chr(12)
			    lgObjRs.MoveNext

		        iDx =  iDx + 1
		        If iDx > C_SHEETMAXROWS_D Then
		           lgStrPrevKey = lgStrPrevKey + 1
		           Exit Do
		        End If               
		    Loop 
			    
		    lgObjRs.Close
	
		End If       

'    End If
    
    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & arrRs(0)       & """" & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & arrRs(1)       & """" & vbCr
    'Response.Write "	.lgPageNo = """ & iIntQueryCount           & """" & vbCr
    'Response.Write "	.lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)          & """" & vbCr
    'Response.Write "	.frm1.hCtrlCd.value =	""" & ConvSPChars(E1_a_ctrl_item(i_a_ctrl_item_ctrl_cd))          & """" & vbCr
    'Response.Write "	.frm1.txtCtrlNM.value = """ & ConvSPChars(E1_a_ctrl_item(i_a_ctrl_item_ctrl_nm))          & """" & vbCr
	
    Response.Write "	.DbQueryOk                                      " & vbCr
    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>                                          " & vbCr

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.DOC_DATE, A.DOC_AMT, A.DEBIT_CREDIT, dbo.ufn_GetCodeName('W1004', A.DEBIT_CREDIT) DEBIT_CREDIT_NM, A.SUMMARY_DESC, A.COMPANY_CD "
            lgStrSQL = lgStrSQL & " , A.COMPANY_NM, A.STOCK_RATE, A.ACQUIRE_AMT, A.COMPANY_TYPE, dbo.ufn_GetCodeName('W1004', A.COMPANY_TYPE) COMPANY_TYPE_NM, A.HOLDING_TERM, A.JUKSU, A.OWN_RGST_NO, A.CO_ADDR "
            lgStrSQL = lgStrSQL & " , A.REPRE_NM, A.STOCK_CNT "

            lgStrSQL = lgStrSQL & " FROM TB_DIVIDEND A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  A.DOC_DATE DESC" & vbcrlf

      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.MINOR_NM, A.MINOR_CD, '' W1, '' W2, A.W3_1, A.W3_2, A.W3_3, A.W3_4, A.W4 "

            lgStrSQL = lgStrSQL & " FROM TB_DIVIDEND_REF A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  A.MINOR_CD ASC" & vbcrlf
    End Select
	PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i, sData

    'On Error Resume Next
    Err.Clear 

	sData = Request("txtSpread")
	PrintLog "1번째 그리드.. : " & sData
	
	If sData <> "" Then
		arrRowVal = Split(sData, gRowSep)                                 '☜: Split Row    data
		lgLngMaxRow = UBound(arrRowVal)
	
		For iDx = 1 To lgLngMaxRow
		    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
		    
		    Select Case arrColVal(0)
		        Case "C"
		                Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
		        Case "U"
		                Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
		        Case "D"
		                Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
		    End Select
		    
		    If lgErrorStatus    = "YES" Then
		       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
		       Exit Sub
		    End If
		    
		Next
	End If
	
	sData = Request("txtSpread2")
	PrintLog "2번째 그리드.. : " & sData
	
	If sData <> "" Then
		' 2번 그리드 
		arrRowVal = Split(sData, gRowSep)                                 '☜: Split Row    data
		lgLngMaxRow = UBound(arrRowVal)
	
		For iDx = 1 To lgLngMaxRow
		    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
		    
		    Select Case arrColVal(0)
		        Case "C"
		                Call SubBizSaveMultiCreate2(arrColVal)                            '☜: Create
		        Case "U"
		                Call SubBizSaveMultiUpdate2(arrColVal)                            '☜: Update
		    End Select
		    
		    If lgErrorStatus    = "YES" Then
		       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
		       Exit For
		    End If
		    
		Next
    End If
End Sub  

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	'On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_DIVIDEND WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, DOC_DATE, DOC_AMT, DEBIT_CREDIT, SUMMARY_DESC " '--, COMPANY_CD" 
	lgStrSQL = lgStrSQL & " , COMPANY_NM, STOCK_RATE, ACQUIRE_AMT, COMPANY_TYPE, HOLDING_TERM, JUKSU, OWN_RGST_NO, CO_ADDR"
	lgStrSQL = lgStrSQL & " , REPRE_NM, STOCK_CNT, INSRT_USER_ID, UPDT_USER_ID )" 
	lgStrSQL = lgStrSQL & " VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"1","D")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_DOC_DATE))),"NULL","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_DOC_AMT), "0"),"0","D")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_DEBIT_CREDIT))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SUMMARY_DESC))),"''","S")     & ","
	'lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_COMPANY_CD))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_COMPANY_NM))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_STOCK_RATE))),"0","D")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_DOC_AMT), "0"),"0","D")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_COMPANY_TYPE))),"''","S")    & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_HOLDING_TERM))),"''","S")    & ","  
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_JUKSU), "0"),"0","D")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_OWN_RGST_NO))),"''","S")    & "," 	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_CO_ADDR))),"''","S")    & "," 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_REPRE_NM))),"''","S")    & "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_STOCK_CNT),0),"0","D")    & "," 

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

Function ParseW3_2(Byval pVal)
	Select Case pVal
		Case "초과"
			ParseW3_2 = ">"
		Case "이상"
			ParseW3_2 = ">="
		Case "이하"
			ParseW3_2 = "<="
		Case "미만"
			ParseW3_2 = "<"
	End Select
End Function

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_DIVIDEND_REF WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, MINOR_NM, MINOR_CD, W3_1, W3_2, W3_3, W3_4, W4 "
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )" 
	lgStrSQL = lgStrSQL & "  VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"1","D")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_MINOR_NM))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_MINOR_CD))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3_1), "0"),"0","D")  & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(ParseW3_2(arrColVal(C_W3_2)))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3_4), "0"),"0","D")  & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(ParseW3_2(arrColVal(C_W3_5)))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")     & ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_DIVIDEND WITH (ROWLOCK) " & vbCrLf
    
    lgStrSQL = lgStrSQL & " SET "  & vbCrLf
    lgStrSQL = lgStrSQL & " DOC_DATE	   = " &  FilterVar(Trim(UCase(arrColVal(C_DOC_DATE))),"NULL","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " DOC_AMT     = " &  FilterVar(UNICDbl(arrColVal(C_DOC_AMT), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " DEBIT_CREDIT       = " &  FilterVar(Trim(UCase(arrColVal(C_DEBIT_CREDIT))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " SUMMARY_DESC     = " &  FilterVar(Trim(UCase(arrColVal(C_SUMMARY_DESC))),"''","S") & "," & vbCrLf
    'lgStrSQL = lgStrSQL & " COMPANY_CD         = " &  FilterVar(Trim(UCase(arrColVal(C_COMPANY_CD))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " COMPANY_NM     = " &  FilterVar(Trim(UCase(arrColVal(C_COMPANY_NM))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " STOCK_RATE = " &  FilterVar(UNICDbl(arrColVal(C_STOCK_RATE), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " ACQUIRE_AMT  = " &  FilterVar(UNICDbl(arrColVal(C_DOC_AMT), "0"),"0","D")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " COMPANY_TYPE       = " &  FilterVar(Trim(UCase(arrColVal(C_COMPANY_TYPE))),"''","S") & ","  & vbCrLf   
    lgStrSQL = lgStrSQL & " HOLDING_TERM      = " &  FilterVar(Trim(UCase(arrColVal(C_HOLDING_TERM))),"","S")   & ","  & vbCrLf    
    lgStrSQL = lgStrSQL & " JUKSU  = " &  FilterVar(UNICDbl(arrColVal(C_JUKSU), "0"),"0","D")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " OWN_RGST_NO      = " &  FilterVar(Trim(UCase(arrColVal(C_OWN_RGST_NO))),"","S")   & ","    & vbCrLf 
    lgStrSQL = lgStrSQL & " CO_ADDR      = " &  FilterVar(Trim(UCase(arrColVal(C_CO_ADDR))),"","S")  & ","      & vbCrLf
    lgStrSQL = lgStrSQL & " REPRE_NM      = " &  FilterVar(Trim(UCase(arrColVal(C_REPRE_NM))),"","S") & ","    & vbCrLf   
    lgStrSQL = lgStrSQL & " STOCK_CNT      = " &  FilterVar(UNICDbl(arrColVal(C_STOCK_CNT), "0"),"0","D")  & ","   & vbCrLf                   
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","            & vbCrLf
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                   & vbCrLf

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  

	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_DIVIDEND_REF WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W3_1	= " &  FilterVar(Trim(UCase(arrColVal(C_W3_1))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W3_2	= " &  FilterVar(Trim(UCase(ParseW3_2(arrColVal(C_W3_2)))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W3_3	= " &  FilterVar(Trim(UCase(arrColVal(C_W3_4))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W3_4	= " &  FilterVar(Trim(UCase(ParseW3_2(arrColVal(C_W3_5)))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & ","

    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  

	PrintLog "SubBizSaveMultiUpdate2 = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_DIVIDEND WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
	
  ' Response.Write lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

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
    On Error Resume Next
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