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
   
	Const BIZ_MNU_ID = "W3107MA1"
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE, sACCT_CD
	Dim iStrData, iStrData2

	Dim C_SEQ_NO1
	Dim C_W1
	Dim C_W1_NM
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7
	Dim C_W8
	Dim C_W9
	Dim C_W10
	Dim C_W11
	
	Dim C_SEQ_NO2
	Dim C_W12
	Dim C_W12_NM
	Dim C_W13
	Dim C_W14
	Dim C_W15
	Dim C_W16
	Dim C_W17
	Dim C_W18
	Dim C_W19

	lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR = Request("txtFISC_YEAR")
    sREP_TYPE = Request("cboREP_TYPE")
    sACCT_CD = Request("txtACCT_CD")

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    	
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
         Case Cstr("MD")
			Call SubBizDeleteMulti()
    End Select

    Call SubCloseDB(lgObjConn)
    
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
	C_SEQ_NO1	= 1	' -- 1번 그리드 
    C_W1		= 2
    C_W1_NM		= 3
    C_W2		= 4
    C_W3		= 5
    C_W4		= 6
    C_W5		= 7
    C_W6		= 8	
    C_W7		= 9
    C_W8		= 10
    C_W9		= 11
    C_W10		= 12	
    C_W11		= 13	
 
 	C_SEQ_NO2	= 1  ' -- 2번 그리드 
    C_W12		= 2 
    C_W12_NM	= 3
    C_W13		= 4
    C_W14		= 5
    C_W15		= 6
    C_W16		= 7
    C_W17		= 8
    C_W18		= 9
    C_W19		= 10
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
    lgStrSQL =            "DELETE TB_BED_DEBT WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_BED_DEBT_CON WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf

	PrintLog "SubBizDelete = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	Call TB_15_DeleData("999999", 0)
 	
End Sub

'========================================================================================
'add 20060123 byHJO
'========================================================================================
Sub SubBizDeleteMulti()
    'On Error Resume Next
    Err.Clear

	' 디테일만 제거한다.

    lgStrSQL = lgStrSQL & "DELETE TB_BED_DEBT_CON WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf

	PrintLog "SubBizDelete = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	'Call TB_15_DeleData("999999", 0)
 	
End Sub
'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iKey4, iIntMaxRows, iLngRow
    Dim iDx
    Dim iLoopMax
    Dim iDataFlag
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    iDataFlag = False

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    iKey4 = FilterVar(sACCT_CD,"''", "S")		' 계정 

     ' 1번째 그리드 
    Call SubMakeSQLStatements("R1",iKey1, iKey2, iKey3, iKey4) 
    
	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

'	    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
'	    Call SetErrorStatus()
	    
	Else
		iDataFlag = True
	    lgstrData = ""
	    
	    iDx = 1
	    
	    Do While Not lgObjRs.EOF
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W1"))			
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W1_NM"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W2"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W3"))			
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W4"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W5"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W6"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W7"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W7_NM"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W8"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W9"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W9_NM"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W10"))			 
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W11"))			 
			iStrData = iStrData & Chr(11) & iDx
			iStrData = iStrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

	        iDx =  iDx + 1
	    Loop 
	    
		lgObjRs.Close
	    
	End If  
	Set lgObjRs = Nothing
	       
    ' 2번째 그리드 
    Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3, iKey4) 
    
	If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

'	    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
'	    Call SetErrorStatus()
	    
	Else
		iDataFlag = True
	    lgstrData2 = ""
	    
	    iDx = 1
	    
	    Do While Not lgObjRs.EOF
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W12"))			
			iStrData2 = iStrData2 & Chr(11) & ""			
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W12_NM"))			
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W13"))	
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W14"))	
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W15"))	
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W16"))		
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W17"))		
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W18"))		
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(lgObjRs("W19"))			 
			iStrData2 = iStrData2 & Chr(11) & iDx
			iStrData2 = iStrData2 & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

	        iDx =  iDx + 1
	    Loop 
	    
		lgObjRs.Close
		Set lgObjRs = Nothing
	    
	End If        
	
	If iDataFlag = False Then
	    'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
	    Call SetErrorStatus()
	End If
    
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
     

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3, pCode4)
    Select Case pMode 
      Case "R1"
			lgStrSQL =			  " SELECT "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W1, A.W1_NM, A.W2, W3, A.W4, A.W5 "
            lgStrSQL = lgStrSQL & " , A.W6, A.W7, dbo.ufn_GetCodeName('W1050', A.W7) W7_NM, A.W8, A.W9, dbo.ufn_GetCodeName('W1050', A.W9) W9_NM, A.W10, A.W11 "

            lgStrSQL = lgStrSQL & " FROM TB_BED_DEBT A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf

			If pCode4 <> "''" Then
			    lgStrSQL = lgStrSQL & "		AND (A.W1 = " & pCode4 	 & vbCrLf
			    lgStrSQL = lgStrSQL & "		OR A.SEQ_NO = '999999') " & vbCrLf
			End If

      Case "R2"
			lgStrSQL =			  " SELECT "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W12, A.W12_NM, A.W13, W14, A.W15, A.W16 "
            lgStrSQL = lgStrSQL & " , A.W17, A.W18, A.W19 "

            lgStrSQL = lgStrSQL & " FROM TB_BED_DEBT_CON A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            
			If pCode4 <> "''" Then
			    lgStrSQL = lgStrSQL & "		AND (A.W12 = " & pCode4 	 & vbCrLf
			    lgStrSQL = lgStrSQL & "		OR A.SEQ_NO = '999999') " & vbCrLf
			End If
    End Select


    lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC " & vbcrlf

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

'	PrintLog "SubBizSaveMulti.."
	
    'On Error Resume Next
    Err.Clear 
    
    ' 헤더 저장 
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
 
'    Select Case lgIntFlgMode
'        Case  OPMD_CMODE                                                             '☜ : Create
'              Call SubBizSaveSingleCreate()  
'        Case  OPMD_UMODE           
'              Call SubBizSaveSingleUpdate()
'    End Select

	PrintLog "1번째 그리드. .: " & Request("txtSpread") 
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
	PrintLog "2번째 그리드.. : " & Request("txtSpread2")
	
	' --- 2번째 그리드 
	arrRowVal = Split(Request("txtSpread2"), gRowSep)                                 '☜: Split Row    data
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
    On Error Resume Next
    Err.Clear

    lgStrSQL =            " INSERT INTO TB_BED_DEBT WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " (CO_CD, FISC_YEAR, REP_TYPE, W1, W2, W3, W4, W5, W6 " 
    lgStrSQL = lgStrSQL & "  , INSRT_USER_ID, UPDT_USER_ID ) " 
    lgStrSQL = lgStrSQL & " VALUES ( " 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","    
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW1"),"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("cboW2"),"''","S") & ","
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
    On Error Resume Next
    Err.Clear

    lgStrSQL =            "UPDATE TB_BED_DEBT WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & "   SET "  & vbCrLf
    lgStrSQL = lgStrSQL & "       W1 = " & FilterVar(Request("txtW1"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W2 = " & FilterVar(Request("cboW2"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W3 = " & FilterVar(Request("txtW3"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W4 = " & FilterVar(Request("txtW4"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W5 = " & FilterVar(Request("txtW5"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W6 = " & FilterVar(Request("txtW6"),"''","S") & "," & vbCrLf

    lgStrSQL = lgStrSQL & "		  UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & "		  UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
 
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
	
	lgStrSQL = "INSERT INTO TB_BED_DEBT WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W1, W1_NM, W2, W3, W4, W5, W6, W7 "  
	lgStrSQL = lgStrSQL & " , W8, W9, W10, W11 " 
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO1), "0"),"1","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1_NM))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W2), ""),"NULL","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W4), ""),"NULL","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S")		& ","


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
	
	lgStrSQL = "INSERT INTO TB_BED_DEBT_CON WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W12, W12_NM, W13, W14, W15, W16, W17, W18, W19 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO2), "0"),"1","D") & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W12))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W12_NM))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W13), ""),"NULL","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W14))),"''","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W15), ""),"NULL","S")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D")		& ","  
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D")		& "," 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W19))),"''","S")		& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate2 = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	If CDbl(arrColVal(C_SEQ_NO2)) = 999999 Then	'W17 계금액은 15호에 밀어준다.
		Call TB_15_PushData(arrColVal(C_SEQ_NO2), arrColVal(C_W17))
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

    lgStrSQL = "UPDATE  TB_BED_DEBT WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W1	= " &  FilterVar(Trim(UCase(arrColVal(C_W1 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W1_NM	= " &  FilterVar(Trim(UCase(arrColVal(C_W1_NM ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W2  = " &  FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W2), ""),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " W3  = " &  FilterVar(Trim(UCase(arrColVal(C_W3 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W4 = " &  FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W4), ""),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " W5 = " &  FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W6 = " &  FilterVar(UNICDbl(arrColVal(C_W6), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W7 = " &  FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W8 = " &  FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W9 = " &  FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W10 = " &  FilterVar(UNICDbl(arrColVal(C_W10), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W11 = " &  FilterVar(Trim(UCase(arrColVal(C_W11))),"''","S") & ","

    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO1), "0"),"''","S")  

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

    lgStrSQL = "UPDATE  TB_BED_DEBT_CON WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W12	= " &  FilterVar(Trim(UCase(arrColVal(C_W12))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W12_NM	= " &  FilterVar(Trim(UCase(arrColVal(C_W12_NM))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W13	= " &  FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W13), ""),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " W14	= " &  FilterVar(Trim(UCase(arrColVal(C_W14))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W15	= " &  FilterVar(UNIConvDateCompanyToDB(arrColVal(C_W15), ""),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " W16 = " &  FilterVar(UNICDbl(arrColVal(C_W16), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W17 = " &  FilterVar(UNICDbl(arrColVal(C_W17), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W18 = " &  FilterVar(UNICDbl(arrColVal(C_W18), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W19	= " &  FilterVar(Trim(UCase(arrColVal(C_W19))),"''","S") & ","
                       
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO2), "0"),"''","S")  

	PrintLog "SubBizSaveMultiUpdate2 = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
	
	If UNICDbl(arrColVal(C_SEQ_NO2), "0") = 999999 Then	'W17 계금액은 15호에 밀어준다.
		Call TB_15_PushData(arrColVal(C_SEQ_NO2), arrColVal(C_W17))
	End If
End Sub

'============================================================================================================
' Name : TB_15_PushData
' Desc : ' (16) 추인금액의 합을 15-2호에 (1)과목명에 대손금 (3)소득처분 유보로 하여 각각 입력하고, 조정내용은 "전기 이전 대손금 부인액을 당기에 손금산입하고 유보처분함."을 입력 
'============================================================================================================
Sub TB_15_PushData(Byval pSeqNo, Byval pW16)
	On Error Resume Next 
	Err.Clear  
	
	lgStrSQL = "EXEC usp_TB_15_PushData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("2")),"''","S") & ", "			' 1호/2호 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("3402")),"''","S") & ", "		' 과목 코드 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pW16, "0"),"0","D")  & ", "			' 금액 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("100")),"''","S") & ", "			' 처분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("전기 이전 대손금 부인액을 당기에 손금산입하고 유보처분함.")),"''","S") & ", "			' 조정내용 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "

	If UNICDbl(pW16, "0") > 0 Then
		PrintLog "TB_15_PushData = " & lgStrSQL
		lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
End Sub

Sub TB_15_DeleData(Byval pSeqNo, Byval pAmt)
	On Error Resume Next 
	Err.Clear  
	
	lgStrSQL = "EXEC usp_TB_15_DeleData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_BED_DEBT WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO1), "0"),"''","S")  

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 2번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_BED_DEBT_CON WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO2), "0"),"''","S")  

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
	If UNICDbl(arrColVal(C_SEQ_NO2), "0") = 999999 Then	'W16 계금액은 15호에서 삭제한다.
		Call TB_15_DeleData(arrColVal(C_SEQ_NO2), arrColVal(C_W16))
	End If

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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
				With parent
					.ggoSpread.Source = .frm1.vspdData
					.ggoSpread.SSShowData "<%=iStrData%>"
					.ggoSpread.Source = .frm1.vspdData2
					.ggoSpread.SSShowData "<%=iStrData2%>"
					.DbQueryOk
				End With
          End If   

       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>C                                                                                                                                                                                                                                                                                                                                                                                                                                        