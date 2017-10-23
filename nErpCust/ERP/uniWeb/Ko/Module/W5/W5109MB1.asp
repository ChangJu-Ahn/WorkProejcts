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
   
	Const BIZ_MNU_ID = "W5109MA1"
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	' -- 그리드 컬럼 정의 
	Dim C_SEQ_NO
	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7
	Dim C_W8
	Dim C_W_DESC

	Dim C_W9_CD
	Dim C_W9_AMT
	Dim C_W9_DESC

	lgErrorStatus   = "NO"
    lgOpModeCRUD    = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

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

	C_SEQ_NO	= 1
	C_W1		= 2
	C_W2		= 3
	C_W3		= 4
	C_W4		= 5
	C_W5		= 6
	C_W6		= 7
	C_W7		= 8
	C_W8		= 9
	C_W_DESC	= 10
	
	C_W9_CD		= 2
	C_W9_AMT	= 3
	C_W9_DESC	= 4
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
    lgStrSQL =            "DELETE TB_22D WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

'printlog "SubMakeSQLStatements = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

    lgStrSQL =            "DELETE TB_22H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

'printlog "SubMakeSQLStatements = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)
 	
 	Call TB_15_DeleData("50", 0 )
 	
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

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
    
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        iIntMaxRows = C_SHEETMAXROWS_D * lgStrPrevKey
        lgstrData2 = ""
        
        iDx = 1
        Do While Not lgObjRs.EOF
        
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W1"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W2"))	
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W3"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W4"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W5"))			
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W6"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W7"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W8"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W_DESC"))
			lgstrData = lgstrData & Chr(11) & iIntMaxRows + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)


		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
        Loop 

	    If iDx <= C_SHEETMAXROWS_D Then
	       lgStrPrevKey = ""
	    End If   

	    Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements
	    lgstrData2 = ""
	
	    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	  
	         lgstrData2 = ""
	        
	    Else
	        
	        iDx = 1
	        Do While Not lgObjRs.EOF
	        
				lgstrData2 = lgstrData2 & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				lgstrData2 = lgstrData2 & Chr(11) & ConvSPChars(lgObjRs("W9_CD"))
				lgstrData2 = lgstrData2 & Chr(11) & ConvSPChars(lgObjRs("W9_AMT"))
				lgstrData2 = lgstrData2 & Chr(11) & ConvSPChars(lgObjRs("W9_DESC"))
				lgstrData2 = lgstrData2 & Chr(11) & iDx
				lgstrData2 = lgstrData2 & Chr(11) & Chr(12)
	
	
			    lgObjRs.MoveNext
	
	            iDx =  iDx + 1
	        Loop 
	    End If
    End If
    

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W1, A.W2, A.W3, A.W4, A.W5, A.W6, A.W7, A.W8, A.W_DESC "
            lgStrSQL = lgStrSQL & " FROM TB_22H A WITH (NOLOCK) " & vbCrLf

            If pCode1 = "''" Then
            	lgStrSQL = lgStrSQL & " WHERE A.CO_CD <> " & pCode1 	 & vbCrLf
            Else
            	lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
            End If

            If pCode2 <> "''" Then
            	lgStrSQL = lgStrSQL & " AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
            End If

            If pCode3 <> "''" Then
            	lgStrSQL = lgStrSQL & " AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            End If

            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC" & vbcrlf
                        
      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.SEQ_NO, A.W9_CD, A.W9_AMT, A.W9_DESC "
            lgStrSQL = lgStrSQL & " FROM TB_22D A WITH (NOLOCK) " & vbCrLf

            If pCode1 = "''" Then
            	lgStrSQL = lgStrSQL & " WHERE A.CO_CD <> " & pCode1 	 & vbCrLf
            Else
            	lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
            End If

            If pCode2 <> "''" Then
            	lgStrSQL = lgStrSQL & " AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
            End If

            If pCode3 <> "''" Then
            	lgStrSQL = lgStrSQL & " AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            End If

            'lgStrSQL = lgStrSQL & " ORDER BY  A.W9_CD ASC" & vbcrlf
    End Select
	'printlog "SubMakeSQLStatements = " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

	'printlog "SubBizSaveMulti.."
	
    On Error Resume Next
    Err.Clear 
    
	'printlog "1번째 그리드. .: " & Request("txtSpread") 
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

	'printlog "2번째 그리드. .: " & Request("txtSpread2") 
	' --- 1번째 그리드 
	arrRowVal = Split(Request("txtSpread2"), gRowSep)                                 '☜: Split Row    data
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
 

End Sub  

     
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : 1번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_22H WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W1, W2, W3, W4, W5, W6, W7, W8, W_DESC "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W4))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W5))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W6))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_DESC))),"''","S")	& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	'printlog "SubBizSaveMultiCreate = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 1번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	
	'On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_22H WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    
	lgStrSQL = lgStrSQL & " W1		= " &  FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")	& ","
	lgStrSQL = lgStrSQL & " W2		= " &  FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S")	& ","
	lgStrSQL = lgStrSQL & " W3		= " &  FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S")	& ","
	lgStrSQL = lgStrSQL & " W4		= " &  FilterVar(Trim(UCase(arrColVal(C_W4))),"''","S")	& ","
	lgStrSQL = lgStrSQL & " W5		= " &  FilterVar(Trim(UCase(arrColVal(C_W5))),"''","S")	& ","
	lgStrSQL = lgStrSQL & " W6		= " &  FilterVar(Trim(UCase(arrColVal(C_W6))),"''","S")	& ","
	lgStrSQL = lgStrSQL & " W7		= " &  FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S")	& ","
	lgStrSQL = lgStrSQL & " W8		= " &  FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & " W_DESC	= " &  FilterVar(Trim(UCase(arrColVal(C_W_DESC))),"''","S")	& ","
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D")	 

	'printlog "SubBizSaveMultiUpdate = " & lgStrSQL
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

    lgStrSQL = "DELETE  TB_22H WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D")	 
	
	'printlog "SubBizSaveMultiDelete = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiCreate2
' Desc : 2번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_22D WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W9_CD, W9_AMT, W9_DESC "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9_CD))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W9_AMT), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9_DESC))),"''","S")	& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	'printlog "SubBizSaveMultiCreate2 = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	If arrColVal(C_W9_CD) = "50" Then	
		If UNICDbl(arrColVal(C_W9_AMT), "0") > 0 Then
			Call TB_15_PushData("50", UNICDbl(arrColVal(C_W9_AMT), "0") )
		Else 
			Call TB_15_DeleData("50", UNICDbl(arrColVal(C_W9_AMT), "0") )
		End If
	End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate2
' Desc : 2번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
	dim i
	
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_22D WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    
	lgStrSQL = lgStrSQL & " W9_CD 	= " &  FilterVar(Trim(UCase(arrColVal(C_W9_CD))),"1","D")  & ", "
	lgStrSQL = lgStrSQL & " W9_AMT	= " &  FilterVar(UNICDbl(arrColVal(C_W9_AMT), "0"),"0","D")		& ", "
	lgStrSQL = lgStrSQL & " W9_DESC	= " &  FilterVar(Trim(UCase(arrColVal(C_W9_DESC))),"''","S")	& ", "
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"1","D")  

	'printlog "SubBizSaveMultiUpdate2 = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf


	' Desc : 	마. 기타 기부금의 합계금액이 있는 경우 15-1호에 (1) 과목 "비지정기부금" (3)소득처분 "기타사외유출"로 입력하고,		
	'			조정내역은 " 비지정기부금을 손금불산입하고 기타사외유출로 처분함."으로 입력함.	
	If arrColVal(C_W9_CD) = "50" Then	
		If UNICDbl(arrColVal(C_W9_AMT), "0") > 0 Then
			Call TB_15_PushData("50", UNICDbl(arrColVal(C_W9_AMT), "0") )
		Else 
			Call TB_15_DeleData("50", UNICDbl(arrColVal(C_W9_AMT), "0") )
		End If
	End If
End Sub

'============================================================================================================
' Name : TB_15_PushData
' Desc : 	마. 기타 기부금의 합계금액이 있는 경우 15-1호에 (1) 과목 "비지정기부금" (3)소득처분 "기타사외유출"로 입력하고,		
'			조정내역은 " 비지정기부금을 손금불산입하고 기타사외유출로 처분함."으로 입력함.		
'============================================================================================================
Sub TB_15_PushData(Byval pSeqNo, Byval pAmt)
	On Error Resume Next 
	Err.Clear  
	
	lgStrSQL = "EXEC usp_TB_15_PushData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("2201")),"''","S") & ", "		' 과목 코드 
	lgStrSQL = lgStrSQL & FilterVar(pAmt,"0","D")  & ", "			' 금액 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("500")),"''","S") & ", "			' 처분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("비지정기부금을 손금불산입하고 기타사외유출로 처분함.")),"''","S") & ", "			' 조정내용 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "

	'printlog "TB_15_PushData = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub

'============================================================================================================
' Name : TB_15_DeleData
' Desc : 	마. 기타 기부금의 합계금액이 있는 경우 15-1호에 (1) 과목 "비지정기부금" (3)소득처분 "기타사외유출"로 입력하고,		
'			조정내역은 " 비지정기부금을 손금불산입하고 기타사외유출로 처분함."으로 입력함.		
'============================================================================================================
Sub TB_15_DeleData(Byval pSeqNo, Byval pAmt)
	On Error Resume Next 
	Err.Clear  
	
	lgStrSQL = "EXEC usp_TB_15_DeleData "
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "

	'printlog "TB_15_DeleData = " & lgStrSQL
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
        Case "SD"
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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .ggoSpread.Source     = .frm1.vspdData2
                .ggoSpread.SSShowData "<%=lgstrData2%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
       
</Script>