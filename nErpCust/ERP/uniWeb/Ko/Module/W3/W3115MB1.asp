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
   
	Const BIZ_MNU_ID = "W3115MA1"
	Const C_SHEETMAXROWS_D = 100
	Dim iStrData, iStrData2, iStrData3
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	' -- 1번 그리드 
	Dim C_SEQ_NO
	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_R1		' 3117의 SEQ_NO
	Dim C_R2		' 3117의 (8)인정이자율 종류 
	Dim C_R3		' 3117의 (5)차감계 (8)이 회사부담이자율인 경우만 
	Dim C_R4		' 3117의 (5)차감계의 합계 (8)이 회사부담이자율인 경우의 값만 합산한 값 
	Dim C_R5		' 3117의 (6)이자수익에 대한 값 
	
	' -- 2번 그리드 
	Dim C_CHILD_SEQ_NO
	Dim C_W6
	Dim C_W7
	Dim C_W8
	Dim C_W9
	Dim C_W10
	
	' -- 3번 그리드 
	Dim C_W_TYPE
	Dim C_SEQ_NO2
	Dim C_W11
	Dim C_W12
	Dim C_W13
	Dim C_W14
	Dim C_W15
	Dim C_W16
	Dim C_W17
	Dim C_W18
	Dim C_W19
	Dim C_W20
	Dim C_W21
	Dim C_W22
	Dim C_W23

	lgErrorStatus   = "NO"
    lgOpModeCRUD    = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    lgPrevNext      = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    lgLngMaxRow     = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	= UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
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
Sub InitSpreadPosVariables()

	'--1번그리드 
	C_SEQ_NO = 1
	C_W1 = 2
	C_W2 = 3
	C_W3 = 4
	C_W4 = 5
	C_W5 = 6
	C_R1 = 7
	C_R2 = 8
	C_R3 = 9
	C_R4 = 10
	C_R5 = 11

	'--2번그리드 
	C_CHILD_SEQ_NO	= 2
	C_W6		= 3
	C_W7		= 4
	C_W8		= 5
	C_W9		= 6
	C_W10		= 7

	'--3번그리드 
	C_W_TYPE = 1
	C_SEQ_NO2 = 2
	C_W11 = 3
	C_W12 = 4
	C_W13 = 5
	C_W14 = 6
	C_W15 = 7
	C_W16 = 8
	C_W17 = 9
	C_W18 = 10
	C_W19 = 11
	C_W20 = 12
	C_W21 = 13
	C_W22 = 14
	C_W23 = 15

	
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
    lgStrSQL =            "DELETE TB_19A_2D WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_19A_2H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf

    lgStrSQL = lgStrSQL & "DELETE TB_19A_1D WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

    lgStrSQL = lgStrSQL & "DELETE TB_19A_1H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf

'PrintLog "SubMakeSQLStatements = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

    Call TB_15_PushData("D")

End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iIntMaxRows, iLngRow
    Dim iDx
    Dim iLoopMax
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R1",iKey1, iKey2, iKey3,"")                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        'Exit Sub
        
    Else
        'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        iStrData = ""
        
        iDx = 1
        
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
			iStrData = iStrData & Chr(11) & iDx
			iStrData = iStrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx = iDx + 1
        Loop 
        
        lgObjRs.Close
        Set lgObjRs = Nothing
 
         ' 2번째 그리드 
        Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3,"") 
        
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
		     lgStrPrevKey = ""
		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		    Call SetErrorStatus()
		    
		Else
		    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    iStrData2 = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
		    
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("CHILD_SEQ_NO"))
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W6"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W7"))	
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W8"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W9"))		
				iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W10"))		
				iStrData2 = iStrData2 & Chr(11) & iDx
				iStrData2 = iStrData2 & Chr(11) & Chr(12)
			    lgObjRs.MoveNext
          		iDx = iDx + 1
		    Loop 
		    
			lgObjRs.Close
			Set lgObjRs = Nothing
		    
		End If  
		       
        ' 3번째 그리드 
        Call SubMakeSQLStatements("R3",iKey1, iKey2, iKey3,"") 
        
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
		     lgStrPrevKey = ""
		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		    Call SetErrorStatus()
		    
		Else
'		    Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData3 = ""
		    
		    iDx = 1
		    
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
				iStrData3 = iStrData3 & Chr(11) & iDx
				iStrData3 = iStrData3 & Chr(11) & Chr(12)
			    lgObjRs.MoveNext
			    iDx = iDx + 1
             
		    Loop 
		    
			lgObjRs.Close
			Set lgObjRs = Nothing
		    
		End If        
    End If
    
     Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
     
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3,pCode4)
    Select Case pMode 
      Case "R1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W1, A.W2, A.W3, A.W4, A.W5, A.R1, A.R2, A.R3, A.R4, A.R5 "
            lgStrSQL = lgStrSQL & " FROM TB_19A_1H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC" & vbcrlf

      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.CHILD_SEQ_NO , A.W6, A.W7, A.W8, A.W9, A.W10 "
            lgStrSQL = lgStrSQL & " FROM TB_19A_1D A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC, A.CHILD_SEQ_NO ASC" & vbcrlf
            
      Case "R3"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W_TYPE, A.SEQ_NO, W11, W12, W13 / 100 AS W13, W14 / 100 AS W14, W15 / 100 AS W15 " & VbCrLf
            lgStrSQL = lgStrSQL & " 	, W16 / 100 AS W16, W17 / 100 AS W17, W18 / 100 AS W18, W19 / 100 AS W19, W20 / 100 AS W20, W21 / 100 AS W21, W22 / 100 AS W22, W23 / 100 AS W23 " & VbCrLf
            lgStrSQL = lgStrSQL & " FROM TB_19A_2H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	UNION " & vbCrLf
			lgStrSQL = lgStrSQL & "	SELECT  "
            lgStrSQL = lgStrSQL & " A.W_TYPE, A.SEQ_NO, A.W11, A.W12, A.W13, A.W14, A.W15, A.W16, A.W17, A.W18, A.W19, A.W20, A.W21, A.W22, A.W23  "
            lgStrSQL = lgStrSQL & " FROM TB_19A_2D A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            lgStrSQL = lgStrSQL & " ORDER BY  A.W_TYPE DESC, A.SEQ_NO ASC" & vbcrlf
      Case "DR1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " ( A.SEQ_NO) " 
            lgStrSQL = lgStrSQL & " FROM TB_19A_1D A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.SEQ_NO = " & pCode4 	 & vbCrLf
            
    End Select
	'PrintLog "SubMakeSQLStatements = " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow,arrColVal2
    Dim iDx , i

	'PrintLog "SubBizSaveMulti.."
	
    On Error Resume Next
    Err.Clear 
    
'    Call SubBizDelete()
    'PrintLog "2번째 그리드.. : " & Request("txtSpread2")
	
	' --- 2번째 그리드 
	arrRowVal = Split(Request("txtSpread2"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal2 = Split(arrRowVal(iDx-1), gColSep)    
              
        Select Case arrColVal2(0)
            Case "C"					
				Call SubBizSaveMultiCreate2(arrColVal2)                            '☜: Create
			
            Case "U"
                    Call SubBizSaveMultiUpdate2(arrColVal2)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete2(arrColVal2)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal2(1) & gColSep
           Exit Sub
        End If
        
    Next
    
	'PrintLog "1번째 그리드. .: " & Request("txtSpread") 
	' --- 1번째 그리드 
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    

        Select Case arrColVal(0)
            Case "C" 
				If arrColVal(C_SEQ_NO)<>"999999" then
					If  fncChkHdr(arrColVal(C_SEQ_NO)) Then
						Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
					Else
						lgErrorStatus="D1"
						'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
						ObjectContext.SetAbort
						Exit Sub
					End IF
				Else
					Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
				End If                    
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit SUB
        End If
        
    Next
 
    
    Call TB_15_PushData("D")
    Call TB_15_PushData("C")
 
 	'PrintLog "3번째 그리드.. : " & Request("txtSpread3")
	
	' --- 3번째 그리드 
	arrRowVal = Split(Request("txtSpread3"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate3(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate3(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete3(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit Sub
        End If
        
    Next   
    

End Sub  

'============================================================================================================
' Name : TB_15_PushData
' Desc : 조정액을 인정이자 프로그램의 (7)구분별로 합계하여  (1)과목명에 "인정이자" (2) 금액에는 구분별 합계금액 
'			(3) 소득처분은 인정이자의 "(7)구분"을 입력하고 
'			조정내용은 " 특수관계자에 대한 가지급금  인정이자를 계산하고 익금산입하고 "소득처분"로 처분함 
'============================================================================================================
Sub TB_15_PushData(ByVal pFlag)
	On Error Resume Next 
	Err.Clear  
	
	lgStrSQL = "EXEC usp_TB_19A_TO_15_PushData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ", "
	lgStrSQL = lgStrSQL & FilterVar(pFlag,"''","S") & " "
	
	PrintLog "usp_TB_19A_TO_15_PushData = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub
    
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : 1번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_19A_1H WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W1, W2, W3, W4, W5, R1, R2, R3, R4, R5 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W2), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W5)), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_R1), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_R2))),"''","S")	& ","	
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_R3), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_R4)), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_R5), "0"),"0","D")		& ","

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
' Name : SubBizSaveCreate2
' Desc : 2번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_19A_1D WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, CHILD_SEQ_NO, W6, W7, W8, W9, W10 "
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D") & ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_CHILD_SEQ_NO), "0"),"1","D") & ","
	If arrColVal(C_W6) <> "계" Then
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W6)), "0"),"0","D")     & "," & vbCrLf
	Else
		lgStrSQL = lgStrSQL & "0," & vbCrLf
	End If
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D")		& ","  
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W9), "0"),"0","D")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")		& "," 
	
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
' Name : SubBizSaveCreate3
' Desc : 3번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate3(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i

	If 	UCase(arrColVal(C_W_TYPE)) = "H" AND UNICDbl(arrColVal(C_SEQ_NO2), "0") = 1 Then
		lgStrSQL = "INSERT INTO TB_19A_2H WITH (ROWLOCK) ("
	Else
		lgStrSQL = "INSERT INTO TB_19A_2D WITH (ROWLOCK) ("
	End If

	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, W_TYPE, SEQ_NO, W11, W12, W13, W14, W15 "  
	lgStrSQL = lgStrSQL & " , W16, W17, W18, W19, W20, W21, W22, W23 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO2), "0"),"1","D") & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W12)), "0"),"0","D")     & "," & vbCrLf

	If 	UCase(arrColVal(C_W_TYPE)) = "H" AND UNICDbl(arrColVal(C_SEQ_NO2), "0") = 1 Then
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W13)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W14)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W15)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W16)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W17)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W18)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W19)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W20)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W21)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W22)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W23)), "0"),"0","D")     & "," & vbCrLf
	Else
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")		& ","  
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")		& ","  
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D")		& ","  
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D")		& ","  
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D")		& "," 
	End If
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate3 = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 1번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i,lgStrSQL
	printlog "insub"
	
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_19A_1H WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_W1 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(UNICDbl(arrColVal(C_W2), "0"),"0","D") & "," 
    lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & ","     
	lgStrSQL = lgStrSQL & " W5		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W5)), "0"),"0","D")     & "," & vbCrLf
	    
    lgStrSQL = lgStrSQL & " R1		= " &  FilterVar(UNICDbl(arrColVal(C_R1), "0"),"0","D") & ","        
    lgStrSQL = lgStrSQL & " R2		= " &  FilterVar(Trim(UCase(arrColVal(C_R2 ))),"''","S") & ","        
    lgStrSQL = lgStrSQL & " R3		= " &  FilterVar(UNICDbl(arrColVal(C_R3), "0"),"0","D") & ","        
    lgStrSQL = lgStrSQL & " R4		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_R4)), "0"),"0","D") & ","
 
    lgStrSQL = lgStrSQL & " R5		= " &  FilterVar(UNICDbl(arrColVal(C_R5), "0"),"0","D") & ","                 
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate2
' Desc : 2번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
	dim i
	
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_19A_1D WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
	If arrColVal(C_W6) <> "계" Then
		lgStrSQL = lgStrSQL & " W6		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W6)), "0"),"0","D")     & "," & vbCrLf
	Else
		lgStrSQL = lgStrSQL & " W6		= 0," & vbCrLf
	End If
	lgStrSQL = lgStrSQL & " W7	= " &  FilterVar(UNICDbl(arrColVal(C_W7), "0"),"0","D") & ","
	lgStrSQL = lgStrSQL & " W8	= " &  FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W9  = " &  FilterVar(UNICDbl(arrColVal(C_W9), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W10 = " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & ","
                         
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
	lgStrSQL = lgStrSQL & "		AND CHILD_SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_CHILD_SEQ_NO))),"0","D") 

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate3
' Desc : 3번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate3(arrColVal)
	dim i
	
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

	If 	UCase(arrColVal(C_W_TYPE)) = "H" AND UNICDbl(arrColVal(C_SEQ_NO2), "0") = 1 Then
	    lgStrSQL = "UPDATE  TB_19A_2H WITH (ROWLOCK) "
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " W11 = " &  FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S")		& ","
		lgStrSQL = lgStrSQL & " W12 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W12)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W13 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W13)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W14 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W14)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W15 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W15)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W16 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W16)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W17 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W17)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W18 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W18)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W19 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W19)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W20 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W20)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W21 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W21)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W22 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W22)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W23 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W23)), "0"),"0","D")     & "," & vbCrLf
	Else
	    lgStrSQL = "UPDATE  TB_19A_2D WITH (ROWLOCK) "
	    lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " W11 = " &  FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S")		& ","
		lgStrSQL = lgStrSQL & " W12 = " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W12)), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & " W13 = " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")		& ","  
		lgStrSQL = lgStrSQL & " W14 = " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & " W15 = " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & " W16 = " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")		& ","  
		lgStrSQL = lgStrSQL & " W17 = " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & " W18 = " &  FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & " W19 = " &  FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D")		& ","  
		lgStrSQL = lgStrSQL & " W20 = " &  FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & " W21 = " &  FilterVar(UNICDbl(arrColVal(C_W21), "0"),"0","D")		& "," 
		lgStrSQL = lgStrSQL & " W22 = " &  FilterVar(UNICDbl(arrColVal(C_W22), "0"),"0","D")		& ","  
		lgStrSQL = lgStrSQL & " W23 = " &  FilterVar(UNICDbl(arrColVal(C_W23), "0"),"0","D")		& "," 
	End If
    
                         
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W_TYPE = " &  FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO2))),"0","D")  

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	PrintLog "SubBizSaveMultiCreate3 = " & lgStrSQL
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 1번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_19A_1H WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
	
	PrintLog "SubBizSaveMultiDelete = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete2
' Desc : 2번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_19A_1D WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  
	lgStrSQL = lgStrSQL & "		AND CHILD_SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_CHILD_SEQ_NO))),"0","D") 
   
	PrintLog "SubBizSaveMultiDelete2 = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete3
' Desc : 3번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete3(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	If 	UCase(arrColVal(C_W_TYPE)) = "H" AND UNICDbl(arrColVal(C_SEQ_NO2), "0") = 1 Then
	    lgStrSQL = "DELETE  TB_19A_2H WITH (ROWLOCK) "
	Else
	    lgStrSQL = "DELETE  TB_19A_2D WITH (ROWLOCK) "
	End If
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W_TYPE = " & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"0","D")  
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO2))),"0","D")  
   
	PrintLog "SubBizSaveMultiCreate3 = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

 Function RemovePercent(Byval pVal)
	RemovePercent = Replace(pVal, "%", "")
End Function

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
'========================================================================================
'grid1 과 grid2는 1:N 관계 ;최소 1건이상 존재해야한다.
'add 20060126 byHJO
'========================================================================================
Function fncChkHDR(strSeqNo)
	Dim iKey1, iKey2, iKey3
	
    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

	fncChkHDR=True

    Call SubMakeSQLStatements("DR1",iKey1, iKey2, iKey3,strSeqNo)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then         
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        fncChkHDR=False        
        Exit function 
    End IF
    
    fncChkHDR=True
End Function         
%>
<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"

		Case "<%=UID_M0001%>"
			If Trim("<%=lgErrorStatus%>") = "NO" Then
          		With parent
					.ggoSpread.Source = .frm1.vspdData
					.ggoSpread.SSShowData "<%=iStrData%>"
					.ggoSpread.Source = .frm1.vspdData2
					.ggoSpread.SSShowData "<%=iStrData2%>"
					.ggoSpread.Source = .frm1.vspdData3
					.ggoSpread.SSShowData "<%=iStrData3%>"
					Call .SetInitGrid3
					.DbQueryOk				
				End With		
			Else
				Call parent.fncNew				
			End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          ElseIf Trim("<%=lgErrorStatus%>")="D1" Then
			Call parent.DBSaveFail			
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>
<%Response.End%>