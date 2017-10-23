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
	Const TYPE_3	= 2		'

	Dim C_W1
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W5
	Dim C_W6
	Dim C_W7
	Dim C_W8

	Dim C_W_TYPE
	Dim C_SEQ_NO
	Dim C_W9
	Dim C_W9_NM
	Dim C_W10
	Dim C_W10_NM
	Dim C_W11
	Dim C_W12
	Dim C_W13
	Dim C_W14
	Dim C_W15
	Dim C_W16
	Dim C_W17
	Dim C_W18
	Dim C_W18_VIEW
	Dim C_W18_VAL
	Dim C_W19
	Dim C_W20
	Dim C_W21
	Dim C_W35

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
	C_W1		= 0	' HTML 인덱스 
	C_W2		= 1
	C_W3		= 2
	C_W4		= 3
	C_W5		= 4
	C_W6		= 5
	C_W7		= 6
	C_W8		= 7

	C_W_TYPE	= 1	' 그리드 인덱스 
	C_SEQ_NO	= 2
	C_W9		= 3
	C_W9_NM		= 4
	C_W10		= 5
	C_W10_NM	= 6
	C_W11		= 7
	C_W12		= 8
	C_W13		= 9
	C_W14		= 10
	C_W15		= 11
	C_W16		= 12
	C_W17		= 13
	C_W18		= 14
	C_W18_VIEW	= 15
	C_W18_VAL	= 16
	C_W19		= 17
	C_W20		= 18
	C_W21		= 19
	C_W35		= 20
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
    lgStrSQL =            "DELETE TB_8_4D WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
 
    lgStrSQL = lgStrSQL & "DELETE TB_8_4H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf

	PrintLog "SubBizDelete = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iStrData2, iIntMaxRows, iLngRow
    Dim iDx, arrRow(2)
    Dim iLoopMax
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        Exit Sub
    Else
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr
		Response.Write "	.IsRunEvents = True " & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.

		For iDx = C_W1 To C_W8	
			lgstrData = lgstrData & "	.frm1.txtData(" & iDx & ").value = """ & lgObjRs(iDx) & """" & vbCrLf
			If Err.number <> 0 Then
				PrintLog "iDx=" & iDx
				Exit Sub
			End If
		Next 
		
		Response.Write lgstrData  &  vbCrLf
		Response.Write "	.IsRunEvents = False " & vbCrLf	' 이벤트가 발생하게 한다.
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>   " & vbCrLf	

        lgObjRs.Close
        Set lgObjRs = Nothing
 
         ' 1번째 그리드 
        Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3) 
        
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then

		    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
				iDx = CDbl(lgObjRs("W_TYPE"))
				
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W_TYPE"))
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W9"))			
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W9_NM"))			
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W10"))	
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W10_NM"))	
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W11"))		
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W12"))		
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W13"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W14"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W15"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W16_1"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W17_1"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W18"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W18_VIEW"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W18_VAL"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W19"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W20"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W21"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W35"))			 
				arrRow(iDx) = arrRow(iDx) & Chr(11) & iIntMaxRows + iLngRow + 1
				arrRow(iDx) = arrRow(iDx) & Chr(11) & Chr(12)

				If CDbl(lgObjRs("SEQ_NO")) <> 999999 Then
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""			
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""		
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""	
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""		
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""	
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""			 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""			 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""			 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W16_2"))			 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ConvSPChars(lgObjRs("W17_2"))			 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""		 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""	 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""	 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""			 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""		 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""		 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & ""		 
					arrRow(iDx) = arrRow(iDx) & Chr(11) & iIntMaxRows + iLngRow + 1
					arrRow(iDx) = arrRow(iDx) & Chr(11) & Chr(12)
				End If
				
			    lgObjRs.MoveNext   			             
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
    Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_2 & ") " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & arrRow(TYPE_2)       & """" & vbCr

    Response.Write "	.ggoSpread.Source = .lgvspdData(" & TYPE_3 & ") " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ &  arrRow(TYPE_3)       & """" & vbCr	
    'Response.Write "	.DbQueryOk                                      " & vbCr
    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>                                          " & vbCr

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R1"
			lgStrSQL =			  " SELECT TOP 1 "
            lgStrSQL = lgStrSQL & " A.W1, A.W2, A.W3, A.W4, A.W5, A.W6, A.W7, A.W8 "
            lgStrSQL = lgStrSQL & " FROM TB_8_4H A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W_TYPE, A.SEQ_NO, A.W9, A.W9_NM, W10, A.W10_NM "
            lgStrSQL = lgStrSQL & " , A.W11, A.W12, A.W13, A.W14, A.W15, A.W16_1, A.W16_2, A.W17_1, A.W17_2, A.W18, A.W18_VIEW, A.W18_VAL "
			lgStrSQL = lgStrSQL & " , A.W19, A.W20, A.W21, A.W35 "
            lgStrSQL = lgStrSQL & " FROM TB_8_4D A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	            
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
    lgIntFlgMode = CDbl(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
	PrintLog "0번째 html..: " & (lgIntFlgMode = OPMD_CMODE) & ";" & Request("txtSpread0") 
	' --- 1번째 그리드 
	arrColVal = Split(Request("txtSpread0") , gColSep)
	
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate(arrColVal)  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate(arrColVal)
    End Select

	PrintLog "1번째 그리드. .: " & Request("txtSpread1") 
	' --- 1번째 그리드 
	arrRowVal = Split(Request("txtSpread1"), gRowSep)                                 '☜: Split Row    data
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
Sub SubBizSaveSingleCreate(arrColVal)
    On Error Resume Next
    Err.Clear

    lgStrSQL =            " INSERT INTO TB_8_4H WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & " (CO_CD, FISC_YEAR, REP_TYPE, W1, W2, W3, W4, W5, W6, W7, W8 "  & vbCrLf
    lgStrSQL = lgStrSQL & "  , INSRT_USER_ID, UPDT_USER_ID ) " & vbCrLf 
    lgStrSQL = lgStrSQL & " VALUES ( " 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","  & vbCrLf   
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W2), "0"),"0","D")		& "," & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D")		& "," & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")		& "," & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D")		& "," & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Replace(arrColVal(C_W6), "%", ""), "0"),"0","D")		& "," & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Replace(arrColVal(C_W7), "%", ""), "0"),"0","D")		& "," & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D")		& "," & vbCrLf
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
Sub SubBizSaveSingleUpdate(arrColVal)
    On Error Resume Next
    Err.Clear

    lgStrSQL =            "UPDATE TB_8_4H WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & "   SET "  & vbCrLf
    lgStrSQL = lgStrSQL & "       W1 = " & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W2 = " & FilterVar(UNICDbl(arrColVal(C_W2), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W3 = " & FilterVar(UNICDbl(arrColVal(C_W3), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W4 = " & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W5 = " & FilterVar(UNICDbl(arrColVal(C_W5), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W6 = " & FilterVar(UNICDbl(replace(arrColVal(C_W6),"%",""), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W7 = " & FilterVar(UNICDbl(REplace(arrColVal(C_W7),"%",""), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W8 = " & FilterVar(UNICDbl(arrColVal(C_W8), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "		  UPDT_DT			= " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & "		  UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD			= " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR	= " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE	= " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf & vbCrLf 
        
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
	
	lgStrSQL = "INSERT INTO TB_8_4D WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, W_TYPE, SEQ_NO, W9, W9_NM, W10, W10_NM, W11, W12, W13 "   & vbCrLf
	lgStrSQL = lgStrSQL & " , W14, W15, W16_1, W17_1, W18, W18_VIEW, W18_VAL, W19, W20, W21, W35, W16_2, W17_2 "   & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W9_NM))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W10))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W10_NM))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W11))),"NULL","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W12))),"NULL","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W16))),"NULL","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W17))),"NULL","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W18))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W18_VIEW))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W18_VAL), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W21)), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(RemovePercent(arrColVal(C_W35)), "0"),"0","D")		& "," & vbCrLf

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W35+2))),"NULL","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W35+3))),"NULL","S")		& "," & vbCrLf

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
' Desc : 1번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_8_4D WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W9		= " &  FilterVar(Trim(UCase(arrColVal(C_W9 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W9_NM	= " &  FilterVar(Trim(UCase(arrColVal(C_W9_NM ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W10		= " &  FilterVar(Trim(UCase(arrColVal(C_W10 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W10_NM	= " &  FilterVar(Trim(UCase(arrColVal(C_W10_NM))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W11		= " &  FilterVar(Trim(UCase(arrColVal(C_W11))),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " W12		= " &  FilterVar(Trim(UCase(arrColVal(C_W12))),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " W13		= " &  FilterVar(UNICDbl(arrColVal(C_W13), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W14		= " &  FilterVar(UNICDbl(arrColVal(C_W14), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W15		= " &  FilterVar(UNICDbl(arrColVal(C_W15), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W16_1	= " &  FilterVar(Trim(UCase(arrColVal(C_W16))),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " W17_1	= " &  FilterVar(Trim(UCase(arrColVal(C_W17))),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " W18		= " &  FilterVar(Trim(UCase(arrColVal(C_W18))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W18_VIEW= " &  FilterVar(Trim(UCase(arrColVal(C_W18_VIEW))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W18_VAL	= " &  FilterVar(UNICDbl(arrColVal(C_W18_VAL), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W19		= " &  FilterVar(UNICDbl(arrColVal(C_W19), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W20		= " &  FilterVar(UNICDbl(arrColVal(C_W20), "0"),"0","D") & ","
    lgStrSQL = lgStrSQL & " W21		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W21)), "0"),"0","D") & ","
    
    
    If arrColVal(C_W_TYPE) = CStr(TYPE_2) Then
		lgStrSQL = lgStrSQL & " W16_2	= " &  FilterVar(Trim(UCase(arrColVal(C_W35+2))),"NULL","S") & ","
		lgStrSQL = lgStrSQL & " W17_2	= " &  FilterVar(Trim(UCase(arrColVal(C_W35+3))),"NULL","S") & ","
    
    Else
		lgStrSQL = lgStrSQL & " W35		= " &  FilterVar(UNICDbl(RemovePercent(arrColVal(C_W35)), "0"),"0","D") & ","
		lgStrSQL = lgStrSQL & " W16_2	= " &  FilterVar(Trim(UCase(arrColVal(C_W35+2))),"NULL","S") & ","
		lgStrSQL = lgStrSQL & " W17_2	= " &  FilterVar(Trim(UCase(arrColVal(C_W35+3))),"NULL","S") & ","

    End If
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND W_TYPE = " & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"0","S")  & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"0","D")  
	
	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	'Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

Function RemovePercent(Byval pVal)
	RemovePercent = Replace(pVal, "%", "")
End Function

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 2번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_8_4D WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W_TYPE = " & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S")
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBQueryOk
          Else
			Parent.FncNew
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
       
</Script>