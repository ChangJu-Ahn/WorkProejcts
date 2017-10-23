<%@ Transaction=required LANGUAGE=VBSCript%>
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

	' -- 그리드 컬럼 정의 
	Dim C_SEQ_NO
	Dim C_W18
	Dim C_W19
	Dim C_W20
	Dim C_W22
	Dim C_W23
	Dim C_W25
	Dim C_W26
	Dim C_W28
	Dim C_W29
	Dim C_W30
	Dim C_W31

	Const BIZ_MNU_ID = "W3125MA1"
	lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR = Request("txtFISC_YEAR")
    sREP_TYPE = Request("cboREP_TYPE")

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
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
	C_SEQ_NO	= 1	' -- 1번 그리드 
    C_W18		= 2
    C_W19		= 3
    C_W20		= 4
    C_W22		= 5
    C_W23		= 6
    C_W25		= 7
    C_W26		= 8
    C_W28		= 9	
    C_W29		= 10	
    C_W30		= 11
    C_W31		= 12
End Sub

'========================================================================================
Sub SubBizQuery()
    'On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizSave()
    'On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizDelete()
    'On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_26AD WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_26AH WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf

'PrintLog "SubBizDelete = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
 	' -- 15에 삭제 
 	Call TB_15_DeleData
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iStrData2, iIntMaxRows, iLngRow
    Dim iDx
    Dim iLoopMax
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("RH",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.txtW1.value = """ & ConvSPChars(lgObjRs("W1"))          & """" & vbCr
		Response.Write "	.frm1.txtW2.value = """ & ConvSPChars(lgObjRs("W2"))          & """" & vbCr
		Response.Write "	.frm1.txtW3.value = """ & ConvSPChars(lgObjRs("W3"))          & """" & vbCr
		Response.Write "	.frm1.txtW4.value = """ & ConvSPChars(lgObjRs("W4"))          & """" & vbCr
		Response.Write "	.frm1.txtW5.value = """ & ConvSPChars(lgObjRs("W5"))          & """" & vbCr
		Response.Write "	.frm1.txtW6.value = """ & ConvSPChars(lgObjRs("W6"))          & """" & vbCr
		Response.Write "	.frm1.txtW7.value = """ & ConvSPChars(lgObjRs("W7"))          & """" & vbCr
		Response.Write "	.frm1.txtW8.value = """ & ConvSPChars(lgObjRs("W8"))          & """" & vbCr
		Response.Write "	.frm1.txtW9.value = """ & ConvSPChars(lgObjRs("W9"))          & """" & vbCr
		Response.Write "	.frm1.txtW10.value = """ & ConvSPChars(lgObjRs("W10"))          & """" & vbCr
		Response.Write "	.frm1.txtW11.value = """ & ConvSPChars(lgObjRs("W11"))          & """" & vbCr
		Response.Write "	.frm1.txtW12.value = """ & ConvSPChars(lgObjRs("W12"))          & """" & vbCr
		Response.Write "	.frm1.txtW13.value = """ & ConvSPChars(lgObjRs("W13"))          & """" & vbCr
		Response.Write "	.frm1.txtW14.value = """ & ConvSPChars(lgObjRs("W14"))          & """" & vbCr
		Response.Write "	.frm1.txtW15.value = """ & ConvSPChars(lgObjRs("W15"))          & """" & vbCr
		Response.Write "	.frm1.txtW16.value = """ & ConvSPChars(lgObjRs("W16"))          & """" & vbCr
		Response.Write "	.frm1.txtW17.value = """ & ConvSPChars(lgObjRs("W17"))          & """" & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr
   
        lgObjRs.Close
        Set lgObjRs = Nothing
 
         ' 1번째 그리드 
        Call SubMakeSQLStatements("RD",iKey1, iKey2, iKey3) 
        
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
		     lgStrPrevKey = ""
		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		    Call SetErrorStatus()
		    
		Else
		    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    'lgstrData = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W18"))			
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W19"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W20"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W22"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W23"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W25"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W26"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W28"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W29"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W30"))			 
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W31"))			 
				
				iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
				iStrData = iStrData & Chr(11) & Chr(12)
			    lgObjRs.MoveNext

		        iDx =  iDx + 1
           
		    Loop 
		    
			lgObjRs.Close
			Set lgObjRs = Nothing
		    
		End If  
		           
    End If
    
     Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
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
      Case "RH"
			lgStrSQL =			  " SELECT TOP 1 "
            lgStrSQL = lgStrSQL & " W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, W17 "
            lgStrSQL = lgStrSQL & " FROM TB_26AH WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "RD"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " SEQ_NO, W18, W19, W20, W22, W23, W25, W26, W28, W29, W30, W31 "
            lgStrSQL = lgStrSQL & " FROM TB_26AD WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
            
    End Select
	PrintLog "SubMakeSQLStatements = " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

	PrintLog "SubBizSaveMulti.."
	
    'On Error Resume Next
    Err.Clear 
    
    ' 헤더 저장 
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
 
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select

	'PrintLog "1번째 그리드. .: " & Request("txtSpread") 
	' --- 1번째 그리드 
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
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
           Exit For
        End If
        
    Next
    
End Sub  

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next
    Err.Clear

    lgStrSQL =            " INSERT INTO TB_26AH WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " (CO_CD, FISC_YEAR, REP_TYPE " 
    lgStrSQL = lgStrSQL & " , W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, W17 " 
    lgStrSQL = lgStrSQL & "  , INSRT_USER_ID, UPDT_USER_ID ) " 
    lgStrSQL = lgStrSQL & " VALUES ( " 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","    
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW2"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW3"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW7"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW9"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW10"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW11"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW12"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW13"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW14"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW15"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW16"), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW17"), "0"),"0","D") & ","
    
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ""
       
    lgStrSQL = lgStrSQL & "   ) " & vbCrLf
    lgStrSQL = lgStrSQL & "EXEC dbo.usp_TB_26A_TO_15_PushData " 
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& " " & vbCrLf 

	PrintLog "SubBizSaveSingleCreate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	
	' -- 15호에 푸쉬 
	Call TB_15_PushData(UNICDbl(Request("txtW9"), 0) + UNICDbl(Request("txtW17"), 0))
End Sub   

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next
    Err.Clear

    lgStrSQL =            "UPDATE TB_26AH WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & "   SET "  & vbCrLf
    lgStrSQL = lgStrSQL & "       W1 = " & FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W2 = " & FilterVar(UNICDbl(Request("txtW2"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W3 = " & FilterVar(UNICDbl(Request("txtW3"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W4 = " & FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W5 = " & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W6 = " & FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W7 = " & FilterVar(UNICDbl(Request("txtW7"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W8 = " & FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W9 = " & FilterVar(UNICDbl(Request("txtW9"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W10 = " & FilterVar(UNICDbl(Request("txtW10"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W11 = " & FilterVar(UNICDbl(Request("txtW11"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W12 = " & FilterVar(UNICDbl(Request("txtW12"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W13 = " & FilterVar(UNICDbl(Request("txtW13"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W14 = " & FilterVar(UNICDbl(Request("txtW14"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W15 = " & FilterVar(UNICDbl(Request("txtW15"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W16 = " & FilterVar(UNICDbl(Request("txtW16"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W17 = " & FilterVar(UNICDbl(Request("txtW17"), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "		  UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & "		  UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
 
    lgStrSQL = lgStrSQL & "EXEC dbo.usp_TB_26A_TO_15_PushData " 
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& " " & vbCrLf 
	
'Response.Write lgstrsql
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
	' -- 15호에 푸쉬 
	Call TB_15_PushData(UNICDbl(Request("txtW9"), 0) + UNICDbl(Request("txtW17"), 0))
End Sub    

     
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : 1번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_26AD WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE "  
	lgStrSQL = lgStrSQL & " , SEQ_NO, W18, W19, W20, W22, W23, W25, W26, W28, W29, W30, W31 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D")		& ","
	If arrColVal(C_W18) <> "계" Then
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")		& ","
	Else
		lgStrSQL = lgStrSQL & "0,"
	End If
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W25), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W30), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D")		& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	'PrintLog "SubBizSaveMultiCreate = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 1번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_26AD WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W18 = " &  FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W19 = " &  FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W20 = " &  FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W22 = " &  FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W23 = " &  FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W25 = " &  FilterVar(UNICDbl(arrColVal(C_W25), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W26 = " &  FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W28 = " &  FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W29 = " &  FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W30 = " &  FilterVar(UNICDbl(arrColVal(C_W30), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W31 = " &  FilterVar(UNICDbl(arrColVal(C_W31), "0"),"0","D") & ","
                   
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 1번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_26AD WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
   
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : 15호서식에 푸쉬 
' Desc :  
'============================================================================================================
Sub TB_15_PushData(Byval pAmt)
	On Error Resume Next 
	Err.Clear  
	
	lgStrSQL = "EXEC usp_TB_15_PushData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(0, "0"),"0","D") & ", "				' 전송자의 순번 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 차/대 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("2601")),"''","S") & ", "		' 과목 코드 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pAmt, "0"),"0","D")  & ", "			' 금액 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("500")),"''","S") & ", "			' 처분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("비업무용 부동산등에 대한 지급이자를 손금불산입하고 기타사외유출 처분함")),"''","S") & ", "			' 조정내용 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & "  "

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
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(0, "0"),"0","D") & ", "		' 전송자의 순번 
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