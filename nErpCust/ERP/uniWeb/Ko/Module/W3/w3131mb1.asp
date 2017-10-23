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


	Const TYPE_1	= 0		' 그리드 배열번호. 
	Const TYPE_2	= 1		
	Const BIZ_MNU_ID = "W3131MA1"
   
	Const C_SHEETMAXROWS_D = 100

	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey, lgCurrGrid,txtSum
			
	Dim C_SEQ_NO
	Dim C_W1
	Dim C_W2
	Dim C_W3_CD
	Dim C_W3
	Dim C_W4_CD
	Dim C_W4
	Dim C_W5
	Dim C_W5_CD
	Dim C_W6

	Dim C_W7
	Dim C_W8
	Dim C_W9
	Dim C_W10
	Dim C_W11
	Dim C_W12
	Dim C_W13
	Dim C_W14
	Dim C_W15
	Dim C_W16
	Dim C_W17
	Dim C_W18

	lgErrorStatus		= "NO"
    lgOpModeCRUD		= Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    txtSum				= Request("txtSum")

    sFISC_YEAR			= Request("txtFISC_YEAR")
    sREP_TYPE			= Request("cboREP_TYPE")
	
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

	C_W1				= 2
    C_W2				= 3
    C_W3_CD				= 4
    C_W3				= 5
    C_W4_CD				= 6
    C_W4				= 7
    C_W5_CD				= 8
    C_W5				= 9
    C_W6				= 10
    
    C_W7				= 2
    C_W8				= 3
    C_W9				= 4
    C_W10				= 5
    C_W11				= 6
    C_W12				= 7
    C_W13				= 8
    C_W14				= 9
    C_W15				= 10
    C_W16				= 11
    C_W17				= 12
    C_W18				= 13
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
    lgStrSQL =            "DELETE TB_39D WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	lgStrSQL = lgStrSQL & "DELETE TB_39H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iIntMaxRows, iLngRow, arrRs(1), sData
    Dim iDx
    Dim iLoopMax, blnData1, blnData2
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	blnData1 = True : blnData2 = True
	
    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

	Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		blnData1 = False

	Else
	    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
	    sData = ""
		    
	    iDx = 1
		    
	    Do While Not lgObjRs.EOF
			sData = sData & "	.Row = " & lgObjRs("SEQ_NO") & vbCrLf
			sData = sData & "	.Col = " & C_SEQ_NO & "	: .Value = """ & lgObjRs("SEQ_NO") & """" & vbCrLf
			sData = sData & "	.Col = " & C_W2		& "	: .Text = """ & lgObjRs("W2") & """" & vbCrLf
			sData = sData & "	.Col = " & C_W3_CD	& "	: .Value = """ & lgObjRs("W3") & """" & vbCrLf
			sData = sData & "	.Col = " & C_W3		& "	: .Text = """ & lgObjRs("W3_NM") & """" & vbCrLf
			sData = sData & "	.Col = " & C_W4_CD	& "	: .Value = """ & lgObjRs("W4") & """" & vbCrLf
			sData = sData & "	.Col = " & C_W4		& "	: .Text = """ & lgObjRs("W4_NM") & """" & vbCrLf
			sData = sData & "	.Col = " & C_W5_CD		& "	: .Value = """ & lgObjRs("W5") & """" & vbCrLf
			sData = sData & "	.Col = " & C_W5 	& "	: .Value = """ & lgObjRs("W5") & """" & vbCrLf
			sData = sData & "	.Col = " & C_W6		& "	: .Value = """ & lgObjRs("W6") & """" & vbCrLf
	
		    lgObjRs.MoveNext

	        iDx =  iDx + 1
	        If iDx > C_SHEETMAXROWS_D Then
	           lgStrPrevKey = lgStrPrevKey + 1
	           Exit Do
	        End If               
	    Loop 
		    
	    lgObjRs.Close
 
 		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent.lgvspdData(" & TYPE_1 & ")  " & vbCr
		Response.Write "	.MaxRows = 6 " & vbCr
		Response.Write "	Call parent.InitSpreadComboBox " & vbCr
		Response.Write "	Call parent.InitData2 " & vbCr
		Response.Write "	Call parent.SetSpreadLock(" & TYPE_1 & ") " & vbCr
		Response.Write sData & vbCrLf
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>   " & vbCr
	End If
	
	' 2번째 그리드 
	Call SubMakeSQLStatements("R1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
	   blnData2 = False
			    
	Else
	    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
	    sData = ""
			    
	    iDx = 1
			    
	    Do While Not lgObjRs.EOF
			sData = sData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			sData = sData & Chr(11) & lgObjRs("W7")	
			sData = sData & Chr(11) & lgObjRs("W8")			
			sData = sData & Chr(11) & lgObjRs("W9")	
			sData = sData & Chr(11) & lgObjRs("W10")		
			sData = sData & Chr(11) & lgObjRs("W11")	
			sData = sData & Chr(11) & lgObjRs("W12")			
			sData = sData & Chr(11) & lgObjRs("W13")			
			sData = sData & Chr(11) & lgObjRs("W14")			
			sData = sData & Chr(11) & lgObjRs("W15")			
			sData = sData & Chr(11) & lgObjRs("W16")			
			sData = sData & Chr(11) & lgObjRs("W17")			
			sData = sData & Chr(11) & lgObjRs("W18")			
			sData = sData & Chr(11) & iIntMaxRows + iLngRow + 1
			sData = sData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

	        iDx =  iDx + 1
	        'If iDx > C_SHEETMAXROWS_D Then
	        '   lgStrPrevKey = lgStrPrevKey + 1
	        '   Exit Do
	        'End If               
	    Loop 
			    
	    lgObjRs.Close
 
 		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_2 & ")  " & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & sData      & """" & vbCr
		Response.Write "	Call .SetTotalRowLine " & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>   " & vbCr
			   
	End If


	If blnData1 = False And blnData2 = False Then
	    'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.FncNew                                      " & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>   " & vbCr
	Else
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.DbQueryOk                                      " & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>   " & vbCr	
			
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
            lgStrSQL = lgStrSQL & "		A.SEQ_NO, A.W2, A.W3, A.W3 W3_NM, A.W4, A.W4 W4_NM, A.W5, A.W6 "
            lgStrSQL = lgStrSQL & " FROM TB_39H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.SEQ_NO BETWEEN 1 AND 4 " & vbCrLf
			lgStrSQL = lgStrSQL & " UNION " & vbcrlf
			lgStrSQL = lgStrSQL & " SELECT " & vbcrlf
            lgStrSQL = lgStrSQL & "		A.SEQ_NO, A.W2, A.W3, dbo.ufn_GetCodeName('W1058', A.W3) W3_NM, A.W4, dbo.ufn_GetCodeName('W1058', A.W4) W4_NM, A.W5, A.W6 "
            lgStrSQL = lgStrSQL & " FROM TB_39H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.SEQ_NO BETWEEN 5 AND 6 " & vbCrLf

	  Case "R1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "		A.SEQ_NO, A.W7, A.W8, A.W9, A.W10, A.W11, A.W12, A.W13, A.W14, A.W15, A.W16, A.W17, A.W18 "
            lgStrSQL = lgStrSQL & " FROM TB_39D A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC" & vbcrlf
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

    On Error Resume Next
    Err.Clear 

	sData = Request("txtSpread0")
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
		    End Select
		   
		    
		    If lgErrorStatus    = "YES" Then
		       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
		       Exit Sub
		    End If
		    
		Next
	End If

	sData = Request("txtSpread1")
	PrintLog "1번째 그리드.. : " & sData
	
	If sData <> "" Then
		arrRowVal = Split(sData, gRowSep)                                 '☜: Split Row    data
		lgLngMaxRow = UBound(arrRowVal)
	
		For iDx = 1 To lgLngMaxRow
		    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
		    
		    Select Case arrColVal(0)
		        Case "C"
		                Call SubBizSaveMultiCreate2(arrColVal)                            '☜: Create
		        Case "U"
		                Call SubBizSaveMultiUpdate2(arrColVal)                            '☜: Update
		        Case "D"
		                Call SubBizSaveMultiDelete2(arrColVal)                            '☜: Delete
		    End Select
		    
		     '** 재고자산금액 합계가 양수이면 				
			'15-2호에 (1)과목 "감가상각비" (2)금액은 동금액을 (3)소득처분은 "유보(감소)"을 입력하고,				
			'조정내역은 " 기업회계기준과 법인세법상 재고자산의 평가금액 차이를 익금산입하고 유보처분함.
		
			If iDx = lgLngMaxRow Then

				If txtSum > 0 Then
					Call TB_15_PushData("1", txtSum, 1, "3901", "400", "기업회계기준과 법인세법상 재고자산의 평가금액 차이를 익금산입하고 유보처분함")
				Else
					Call TB_15_PushData("2", txtSum, 1, "3901", "100", "기업회계기준과 법인세법상 재고자산의 평가금액 차이를 손금산입하고 유보처분함")
				End If
			End If
		    
		    
		    
		    If lgErrorStatus    = "YES" Then
		       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
		       Exit Sub
		    End If
		    
		Next
	End If	
End Sub  

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_39H WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE " 
	lgStrSQL = lgStrSQL & " , SEQ_NO, W2, W3, W4, W5, W6"
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )" 
	lgStrSQL = lgStrSQL & " VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"NULL","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W3_CD))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W4_CD))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W5_CD))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_W6)),"''","S")     & ","
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	lgStrSQL = "INSERT INTO TB_39D WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE " 
	lgStrSQL = lgStrSQL & " , SEQ_NO, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, W17, W18 "
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )" 
	lgStrSQL = lgStrSQL & " VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_W7)),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_W8)),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_W9)),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W10))),"''","S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")     & "," & vbCrLf

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_39H WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 

    lgStrSQL = lgStrSQL & " W2     = " &  FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W3     = " &  FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W4     = " &  FilterVar(Trim(UCase(arrColVal(C_W4))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W5     = " &  FilterVar(Trim(UCase(arrColVal(C_W5_CD))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W6     = " &  FilterVar(Trim(arrColVal(C_W6)),"''","S") & ","

    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf  

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

    lgStrSQL = "UPDATE  TB_39D WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 

    lgStrSQL = lgStrSQL & " W7     = " &  FilterVar(Trim(arrColVal(C_W7)),"''","S") & ","
    lgStrSQL = lgStrSQL & " W8     = " &  FilterVar(Trim(arrColVal(C_W8)),"''","S") & ","
    lgStrSQL = lgStrSQL & " W9     = " &  FilterVar(Trim(arrColVal(C_W9)),"''","S") & ","
    lgStrSQL = lgStrSQL & " W10     = " &  FilterVar(Trim(UCase(arrColVal(C_W10))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W11     = " &  FilterVar(UNICDbl(arrColVal(C_W11), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W12     = " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W13     = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W14     = " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W15     = " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W16     = " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W17     = " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W18     = " &  FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D") & ","

    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S") 	 & vbCrLf  

	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
	
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_39D WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"''","S")  
	
  ' Response.Write lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
	Call TB_15_DeleData()
	
End Sub

'============================================================================================================
' Name : 15호서식에 푸쉬 
' Desc :  
'============================================================================================================
Sub TB_15_PushData(Byval pType, Byval pAmt, Byval pSeqNo, Byval pAcctCd, Byval pCode, Byval pDesc)
	On Error Resume Next 
	Err.Clear  

	lgStrSQL = "EXEC usp_TB_15_PushData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "				' 전송자의 순번 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pType)),"''","S") & ", "			' 차/대 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pAcctCd)),"''","S") & ", "		' 과목 코드 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pAmt, "0"),"0","D")  & ", "			' 금액 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pCode)),"''","S") & ", "			' 처분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(pDesc)),"''","S") & ", "			' 조정내용 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "
	
		
	PrintLog "TB_15_PushData = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub

Sub TB_15_DeleData()
	On Error Resume Next 
	Err.Clear  

	lgStrSQL = "EXEC usp_TB_15_DeleData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(-1, "0"),"0","D") & ", "		' 전송자의 순번 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "

	PrintLog "TB_15_DeleData = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
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