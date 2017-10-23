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
	Const BIZ_MNU_ID = "W2101MA1"
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
    lgStrSQL =            "DELETE TB_16_2D1 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_16_2D2 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf

    lgStrSQL = lgStrSQL & "DELETE TB_16_2H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	
	PrintLog "SubMakeSQLStatements = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
 	Call TB_15_DeleData()
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

    Call SubMakeSQLStatements("R1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
        'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.txtW1.value = """ & ConvSPChars(lgObjRs("W1"))          & """" & vbCr
		Response.Write "	.frm1.cboW2.value = """ & ConvSPChars(lgObjRs("W2"))          & """" & vbCr
		Response.Write "	.frm1.txtW3.value = """ & ConvSPChars(lgObjRs("W3"))          & """" & vbCr
		Response.Write "	.frm1.txtW4.value = """ & ConvSPChars(lgObjRs("W4"))          & """" & vbCr
		Response.Write "	.frm1.txtW5.value = """ & ConvSPChars(lgObjRs("W5"))          & """" & vbCr
		Response.Write "	.frm1.txtW6.value = """ & ConvSPChars(lgObjRs("W6"))          & """" & vbCr
		If lgObjRs("W2") <> "" Then
			Response.Write "	Call .cboW2_onChange" & vbCr
		End If
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr
   
		'iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W1"))			
		'iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W2"))	
		'iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W2_NM"))	
		'iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W3"))	
		'iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W4"))		
		'iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W5"))		
		'iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W6"))		 
		'iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
		'iStrData = iStrData & Chr(11) & Chr(12)

        lgObjRs.Close
        Set lgObjRs = Nothing
 
         ' 1번째 그리드 
        Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3) 
        
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then

		    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W7"))			
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W8"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W8_NM"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W9"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W10"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W11"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W12"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W13"))			 
				iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
				iStrData = iStrData & Chr(11) & Chr(12)
			    lgObjRs.MoveNext

		        iDx =  iDx + 1
		        If iDx > C_SHEETMAXROWS_D Then
		           lgStrPrevKey = lgStrPrevKey + 1
		           Exit Do
		        End If               
		    Loop 
		    
			lgObjRs.Close
			Set lgObjRs = Nothing
		    
		End If  
		       
        ' 2번째 그리드 
        Call SubMakeSQLStatements("R3",iKey1, iKey2, iKey3) 
        
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
  
		    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData2 = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W14"))			
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W15"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W16"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W17"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W18"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W19"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W20"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W21"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("HEAD_SEQ_NO"))	
				iStrData2 = iStrData2 & Chr(11) & ""		 
				iStrData2 = iStrData2 & Chr(11) & iIntMaxRows + iLngRow + 1
				iStrData2 = iStrData2 & Chr(11) & Chr(12)
			    lgObjRs.MoveNext

		        iDx =  iDx + 1
		        'If iDx > C_SHEETMAXROWS_D Then
		        '   lgStrPrevKey = lgStrPrevKey + 1
		        '   Exit Do
		        'End If               
		    Loop 
		    
			lgObjRs.Close
			Set lgObjRs = Nothing
		    
		End If        
    End If
    
     Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     
    'If iDx <= C_SHEETMAXROWS_D Then
    '   lgStrPrevKey = ""
    'End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_1 & ") " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr

    Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_2 & ") " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData2       & """" & vbCr	
    Response.Write "	.DbQueryOk                                      " & vbCr
    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>                                          " & vbCr

End Sub

Function GetW16(Byval pIndex)
	Select Case pIndex
		Case "0"
			GetW16 = "100"
		Case "1"
			GetW16 = "90"
		Case "2"
			GetW16 = "60"
		Case "3"
			GetW16 = "50"
		Case "4"
			GetW16 = "30"
	End Select
End Function

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R1"
			lgStrSQL =			  " SELECT TOP 1 "
            lgStrSQL = lgStrSQL & " A.W1, A.W2, W2_NM, A.W3, A.W4 "
            lgStrSQL = lgStrSQL & " , A.W5, A.W6 "

            lgStrSQL = lgStrSQL & " FROM TB_16_2H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            'lgStrSQL = lgStrSQL & " ORDER BY  A.W2 ASC" & vbcrlf

      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W7, A.W8, W8_NM, A.W9, A.W10 "
            lgStrSQL = lgStrSQL & " , A.W11, A.W12, A.W13 "

            lgStrSQL = lgStrSQL & " FROM TB_16_2D1 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            'lgStrSQL = lgStrSQL & " ORDER BY  A.W7 ASC" & vbcrlf	' 법인명 순서 
            
      Case "R3"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W14, A.W15, A.W16, A.W17, A.W18 "
            lgStrSQL = lgStrSQL & " , A.W19, A.W20, A.W21, HEAD_SEQ_NO "

            lgStrSQL = lgStrSQL & " FROM TB_16_2D2 A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            'lgStrSQL = lgStrSQL & " ORDER BY  A.W14 ASC" & vbcrlf
            
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
	
    On Error Resume Next
    Err.Clear 
    
    ' 헤더 저장 
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
 
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select

	PrintLog "1번째 그리드. .: " & Request("txtSpread0") 
	' --- 1번째 그리드 
	arrRowVal = Split(Request("txtSpread0"), gRowSep)                                 '☜: Split Row    data
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
	PrintLog "2번째 그리드.. : " & Request("txtSpread1")
	
	' --- 2번째 그리드 
	arrRowVal = Split(Request("txtSpread1"), gRowSep)                                 '☜: Split Row    data
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
    'On Error Resume Next
    Err.Clear

    lgStrSQL =            " INSERT INTO TB_16_2H WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " (CO_CD, FISC_YEAR, REP_TYPE, W1, W2, W2_NM, W3, W4, W5, W6 " 
    lgStrSQL = lgStrSQL & "  , INSRT_USER_ID, UPDT_USER_ID ) " 
    lgStrSQL = lgStrSQL & " VALUES ( " 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","    
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW1"),"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("cboW2"),"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW2_NM"),"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW3"),"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW4"),"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW5"),"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW6"),"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ""
       
    lgStrSQL = lgStrSQL & "   ) " 

	PrintLog "SubBizSaveSingleCreate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub   

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    'On Error Resume Next
    Err.Clear

    lgStrSQL =            "UPDATE TB_16_2H WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & "   SET "  & vbCrLf
    lgStrSQL = lgStrSQL & "       W1 = " & FilterVar(Request("txtW1"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W2 = " & FilterVar(Request("cboW2"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W2_NM = " & FilterVar(Request("txtW2_NM"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W3 = " & FilterVar(Request("txtW3"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W4 = " & FilterVar(Request("txtW4"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W5 = " & FilterVar(Request("txtW5"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W6 = " & FilterVar(Request("txtW6"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "		  UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & "		  UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf & vbCrLf 
        
	PrintLog "SubBizSaveSingleUpdate = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub    

     
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : 1번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_16_2D1 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W7, W8, W8_NM, W9, W10, W11, W12, W13 "   & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO1), "0"),"1","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W8))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W8_NM))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W10))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")		& "," & vbCrLf

	'lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_DOC_AMT), "0"),"0","D")     & ","

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
' Desc : 2번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_16_2D2 WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W14, W15, W16, W17, W18, W19, W20, W21, HEAD_SEQ_NO "   & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES ("  & vbCrLf

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO2), "0"),"1","D") & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W14))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")		& ","   & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")		& ","   & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D")		& ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")		& ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D")		& ","   & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D")		& ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D")		& ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_HEAD_SEQ_NO1), "0"),"1","D") & "," & vbCrLf
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	If CDbl(arrColVal(C_SEQ_NO2)) = 999999 Then	'W21 계금액은 15호에 밀어준다.
		Call TB_15_PushData(arrColVal(C_SEQ_NO2), arrColVal(C_W21))
	End If

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 1번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_16_2D1 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W7	= " &  FilterVar(Trim(UCase(arrColVal(C_W7 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W8  = " &  FilterVar(Trim(UCase(arrColVal(C_W8 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W9  = " &  FilterVar(Trim(UCase(arrColVal(C_W9 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W10 = " &  FilterVar(Trim(UCase(arrColVal(C_W10))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W11 = " &  FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W12 = " &  FilterVar(UNICDbl(arrColVal(C_W12), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W13 = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & ","
                   
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO1))),"0","D")  

	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 2번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
	dim i
	
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_16_2D2 WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W14	= " &  FilterVar(Trim(UCase(arrColVal(C_W14))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W15 = " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W16 = " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W17 = " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W18 = " &  FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W19 = " &  FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W20 = " &  FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W21 = " &  FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D") & ","
                       
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO2))),"0","D")  

	PrintLog "SubBizSaveMultiUpdate2 = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
	If CDbl(arrColVal(C_SEQ_NO2)) = 999999 Then	'W21 계금액은 15호에 밀어준다.
		Call TB_15_PushData(arrColVal(C_SEQ_NO2), arrColVal(C_W21))
	End If

End Sub

Sub TB_15_PushData(Byval pSeqNo, Byval pW21)
	On Error Resume Next 
	Err.Clear  
	
	lgStrSQL = "EXEC usp_TB_15_PushData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("2")),"''","S") & ", "			' 차/대 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1601")),"''","S") & ", "		' 과목 코드 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pW21, "0"),"0","D")  & ", "			' 금액 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("200")),"''","S") & ", "			' 처분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("수입배당금을 익금불산입하고 기타 처분함.")),"''","S") & ", "			' 조정내용 
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
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(999999, "0"),"0","D") & ", "		' 전송자의 순번 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("2")),"''","S") & ", "			' 1호/2호 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "


	PrintLog "TB_15_DeleData = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 1번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_16_2D2 WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
   
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 2번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_16_2D2 WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
   
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