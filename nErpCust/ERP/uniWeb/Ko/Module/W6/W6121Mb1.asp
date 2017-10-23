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
	Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.
	
	' -- 그리드 컬럼 정의 
    Dim C_W101			    ' (101)구분코드 
	Dim C_W101_Nm		    ' 구분명 
	Dim C_Law	    	    ' 근거법 조항 
	Dim C_W102			    ' (102)계산기준 
    Dim C_CODE			    ' 코드 
	Dim C_W103_AMT	        ' 투자(지출)금액 
	Dim C_W103_RATE_VAL	    ' 코드명 
	Dim C_W103_RATE		    ' 공제율 
	Dim C_W103		        ' 계산내역	
	Dim C_W104			    ' 공제세액 
	Dim C_Limit_RATE	    ' 한도율 
	Dim C_Limit_AMT		    ' 한도금액 
	
	Dim C_SEQ_NO        
	Dim C_W105          	   '(105) 구분	
	Dim C_W105_Nm          	   '(105) 구분	
	Dim C_W106		           '(106) 사업년도	
	Dim C_W107		    	   '(107) 당기분 
	Dim C_W108		    	   '(108) 이월분 
	Dim C_W109		    	   '(109) 당기분 
	Dim C_W110		    	   '(110) 1차년도 
	Dim C_W111		    	   '(111) 2차년도 
	Dim C_W112		    	   '(112) 3차년도 
	Dim C_W113		    	   '(113) 4차년도 
	Dim C_W114		    	   '(114) 계 
	Dim C_W115		    	   '(115) 최저한 세적용에 다른 미공제액 
	Dim C_W116		    	   '(116) 공제세액(114-115)
	Dim C_W117		    	   '(117) 소멸 
	Dim C_W118		    	   '(118) 이월액(107 + 108 + 116 - 117)



	lgErrorStatus    = "NO"
    lgOpModeCRUD     = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    'lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    'lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    'lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
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
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
	C_W101			= 1   ' (101)구분코드 
	C_W101_Nm		= 2   ' 구분명 
	C_Law	    	= 3   ' 근거법 조항 
	C_W102			= 4   ' (102)계산기준 
    C_CODE			= 5   ' 코드 
	C_W103_AMT	    = 6   ' 투자(지출)금액 
	C_W103_RATE_VAL	= 7   ' 코드명 
	C_W103_RATE		= 8   ' 공제율	
	C_W103			= 9   ' 공제세액 
	C_W104			= 10   ' 공제세액 
	C_Limit_RATE	= 11  ' 한도율 
	C_Limit_AMT		= 12  ' 한도금액 
	
	C_SEQ_NO        =1
	C_W105          =2	   '(105) 구분	
	C_W105_NM       =3	   '(105) 구분명	
	C_W106		    =4     '(106) 사업년도	
	C_W107		    =5	   '(107) 당기분 
	C_W108		    =6	   '(108) 이월분 
	C_W109		    =7	   '(109) 당기분 
	C_W110		    =8	   '(110) 1차년도 
	C_W111		    =9	   '(111) 2차년도 
	C_W112		    =10	   '(112) 3차년도 
	C_W113		    =11	   '(113) 4차년도 
	C_W114		    =12	   '(114) 계 
	C_W115		    =13	   '(115) 최저한 세적용에 다른 미공제액 
	C_W116		    =14	   '(116) 공제세액(114-115)
	C_W117		    =15	   '(117) 소멸 
	C_W118		    =16	   '(118) 이월액(107 + 108 + 116 - 117)
	
	
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
    'On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_8_3_A WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

	' 디테일부터 제거한다.
    lgStrSQL = lgStrSQL & "DELETE TB_8_3_B WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf  
	
	PrintLog "SubBizDelete = " & lgStrSQL 
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(1), iIntMaxRows, iLngRow,iLngCol
    Dim iRow, iKey1, iKey2, iKey3
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
       ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = "" : iLngRow = 1
        
		iRow = 0
		arrRs(iRow) = ""
		
		iLngCol = lgObjRs.Fields.Count
		
		iDx = 1

		
		        lgstrData = lgstrData & " With parent.lgvspdData(" & TYPE_1 & ")" & vbCr
				lgstrData = lgstrData & "	.Redraw = false " & vbCr
				Do While Not lgObjRs.EOF
					lgstrData = lgstrData & "	.Row = " &iDx & "" & vbCrLf
					lgstrData = lgstrData & "	.Col = 0 : .value = """" " & vbCrLf
				    lgstrData = lgstrData & "	.Col = " & C_CODE          & " : .value   = """ & ConvSPChars(lgObjRs("W_Code"))    & """" & vbCrLf
				    
				    ' -- 저장후 코드13(사용자정의)에 데이타 미 출력 수정: 2006.02.24
				    If lgObjRs("W_Code") = "13" Then
						lgstrData = lgstrData & "	.Col = " & C_W101_Nm          & " : .value   = """ & ConvSPChars(lgObjRs("W101_NM"))    & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_Law          & " : .value   = """ & ConvSPChars(lgObjRs("W_Law"))    & """" & vbCrLf
						lgstrData = lgstrData & "	.Col = " & C_W102          & " : .value   = """ & ConvSPChars(lgObjRs("W102"))    & """" & vbCrLf
					End If
					
					lgstrData = lgstrData & "	.Col = " & C_W103_AMT      & " : .value   = """ & ConvSPChars(lgObjRs("C_W103_AMT")) & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W103_RATE_VAL & " : .text   = """ & ConvSPChars(lgObjRs("C_W103_RATE_VAL")) & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W103_RATE     & " : .text   = """ & ConvSPChars(lgObjRs("C_W103_RATE")) & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W103		   & " : .value   = """ & ConvSPChars(lgObjRs("C_W103"))& """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W104		   & " : .value   = """ & ConvSPChars(lgObjRs("C_W104")) & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_Limit_RATE	   & " : .value   = """ & ConvSPChars(lgObjRs("C_Limit_RATE")) & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_Limit_AMT	   & " : .value   = """ & ConvSPChars(lgObjRs("C_Limit_AMT"))& """" & vbCrLf
                    lgstrData = lgstrData & "	.Col = " & C_Limit_AMT + 1 & " : .value =    """  &  iDx  & """" & vbCrLf

				If Err.number <> 0 Then
					PrintLog "iDx=" & iDx
					Exit Sub
				End If
		        iDx = iDx +1    
				lgObjRs.MoveNext
		
			Loop
			
			lgstrData = lgstrData & "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE" & vbCrLf
			lgstrData = lgstrData & "	.Redraw = True " & vbCr
	        lgstrData = lgstrData & "	 End With " & vbCr

		   iLngRow = 1
		   lgObjRs.Close
		
		
	If lgstrData <> "" Then	
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write lgstrData
			Response.Write " </Script>          " & vbCr
	End If
	       
	
	
	END IF 	
		
		Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
		     lgStrPrevKey = ""
		    'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		    Call SetErrorStatus()
		    
		Else
		   ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = "" : iLngRow = 1
		    

			
			'**************8
							 ' Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
					    lgstrData = "" : iLngRow = 1
					    
						iRow = 0
						arrRs(iRow) = ""
						
						iLngCol = lgObjRs.Fields.Count
						
						iDx = 1
            dim strGubn

						
						        lgstrData = lgstrData & " With parent.lgvspdData(" & TYPE_2 & ")" & vbCr
								lgstrData = lgstrData & "	.Redraw = false " & vbCr
								
								
								Do While Not lgObjRs.EOF
								  
								         
								            lgstrData = lgstrData & "  parent.ggoSpread.InsertRow " &iDx & "  ,1  " & vbCrLf
								          
											lgstrData = lgstrData & "	.Row = " &iDx & "" & vbCrLf
											
								            if   strGubn =   ConvSPChars(lgObjRs("W105") ) or   right( ConvSPChars(lgObjRs("W105") ) ,2 ) ="99" then
						                         lgstrData = lgstrData & "  .Col = " &  C_W105 & ": .CellType = 1    " & vbCrLf
												 lgstrData = lgstrData & "  .Col = " &  C_W105_NM & ": .CellType = 1 : .TypeMaxEditLen  = 100 " & vbCrLf
												 lgstrData = lgstrData & "  .Col = " &  C_W106 & ": .CellType = 1 : .TypeMaxEditLen  = 100 : .TypeHAlign = 2 " & vbCrLf
									
                                            end if
								          
											lgstrData = lgstrData & "	.Col = 0 : .value = """" " & vbCrLf
											
											
											if  right( ConvSPChars(lgObjRs("W105") ) ,2 ) ="99" then
											    lgstrData = lgstrData & "	.Col = " & C_SEQ_NO          & " : .value   = """ & 999999   & """" & vbCrLf
											    lgstrData = lgstrData & "	.Col = " & C_W105			 & " : .value   = """ & 99 & """" & vbCrLf
											else
											    lgstrData = lgstrData & "	.Col = " & C_SEQ_NO          & " : .value   = """ &  UNIConvNum(lgObjRs("SEQ_NO"), 0)   & """" & vbCrLf
											    lgstrData = lgstrData & "	.Col = " & C_W105			 & " : .text   = """ & ConvSPChars(lgObjRs("W105")) & """" & vbCrLf
											end if    
											
											lgstrData = lgstrData & "	.Col = " & C_W105_NM		 & " : .text   = """ & ConvSPChars(lgObjRs("W105_NM")) & """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W106			 & " : .text   = """ & ConvSPChars(lgObjRs("W106"))  & """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W107			 & " : .value   = """ & ConvSPChars(lgObjRs("C_W107")) & """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W108	   & " : .value   = """ &  UNIConvNum(lgObjRs("C_W108"), 0)& """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W109	   & " : .value   = """ & UNIConvNum(lgObjRs("C_W109"), 0)& """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W110	   & " : .value   = """ & UNIConvNum(lgObjRs("C_W110"), 0)& """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W111	   & " : .value   = """ & UNIConvNum(lgObjRs("C_W111"), 0)& """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W112	   & " : .value   = """ & UNIConvNum(lgObjRs("C_W112"), 0)& """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W113	   & " : .value   = """ & UNIConvNum(lgObjRs("C_W113"), 0)& """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W114	   & " : .value   = """ & UNIConvNum(lgObjRs("C_W114"), 0)& """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W115	   & " : .value   = """ & UNIConvNum(lgObjRs("C_W115"), 0)& """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W116	   & " : .value   = """ & UNIConvNum(lgObjRs("C_W116"), 0)& """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W117	   & " : .value   = """ & UNIConvNum(lgObjRs("C_W117"), 0)& """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W118	   & " : .value   = """ & UNIConvNum(lgObjRs("C_W118"), 0)& """" & vbCrLf
									
											lgstrData = lgstrData & "	.Col = " & C_W118 + 1 & " : .value =    """  &  iDx  & """" & vbCrLf
									        
                                            strGubn =  ConvSPChars(lgObjRs("W105"))
								If Err.number <> 0 Then
									PrintLog "iDx=" & iDx
									Exit Sub
								End If
						        iDx = iDx +1    
								lgObjRs.MoveNext
								
						         
							Loop
							
                            
							
							lgstrData = lgstrData & "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE" & vbCrLf
							lgstrData = lgstrData & "	.Redraw = True " & vbCr
					        lgstrData = lgstrData & "	 End With " & vbCr

						   iLngRow = 1
						   lgObjRs.Close
						
						
					If lgstrData <> "" Then	
							Response.Write " <Script Language=vbscript>	                        " & vbCr
							Response.Write lgstrData
							Response.Write " </Script>          " & vbCr
					End If
	
	
	
		End If
		   
		   
 
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	
	Response.Write " With parent                                        " & vbCr

	'If lgErrorStatus = "NO" Then
		Response.Write "	.DbQueryOk                                      " & vbCr
	'End If
	
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
            lgStrSQL = lgStrSQL & " A.W101, A.W101_NM, A.W_Law, A.W102, A.W_Code, A.C_W103_AMT ,"
            lgStrSQL = lgStrSQL & " A.C_W103_RATE_VAL, A.C_W103_RATE, A.C_W103, A.C_W104, A.C_Limit_RATE, A.C_Limit_AMT "
            lgStrSQL = lgStrSQL & " FROM TB_8_3_A A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
            lgStrSQL = lgStrSQL & " ORDER BY  CAST(A.W101 AS int) ASC" & vbcrlf
            
      Case "R2"
      
            lgStrSQL =			  "SELECT "
            lgStrSQL = lgStrSQL & "  SEQ_NO   , A.W105, A.W105_NM, A.W106, A.C_W107, A.C_W108, A.C_W109, A.C_W110, A.C_W111, A.C_W112, A.C_W113, A.C_W114, A.C_W115, A.C_W116, A.C_W117, A.C_W118 "
			lgStrSQL = lgStrSQL & " FROM TB_8_3_B A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "  union all"
			lgStrSQL = lgStrSQL & "  SELECT 9999999 SEQ_NO ,  A.W105 +'Sum99' , '','소계',Sum(A.C_W107), Sum(A.C_W108), Sum(A.C_W109), Sum(A.C_W110), Sum(A.C_W111), Sum(A.C_W112), Sum(A.C_W113), Sum(A.C_W114), Sum(A.C_W115), Sum(A.C_W116),Sum(A.C_W117), Sum(A.C_W118 )"
			lgStrSQL = lgStrSQL & "  FROM TB_8_3_B A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & "  WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "  GROUP BY   A.W105,  A.W105_NM " & vbcrlf
			lgStrSQL = lgStrSQL & "	 union all"
			lgStrSQL = lgStrSQL & "	 SELECT 9999999 SEQ_NO ,  '99999' , '합계','',Sum(A.C_W107), Sum(A.C_W108), Sum(A.C_W109), Sum(A.C_W110), Sum(A.C_W111), Sum(A.C_W112), Sum(A.C_W113), Sum(A.C_W114), Sum(A.C_W115), Sum(A.C_W116),Sum(A.C_W117), Sum(A.C_W118 )"
			lgStrSQL = lgStrSQL & "  FROM TB_8_3_B A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & "  WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		having Sum(A.C_W107) > 0 " & vbCrLf
			lgStrSQL = lgStrSQL & "  ORDER BY  W105 ,  W106  ASC " & vbcrlf
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

    'On Error Resume Next
    Err.Clear 

	' 디테일 삭제 
	                           '☜: Delete
	
	PrintLog "txtSpread = " & Request("txtSpread")
	
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
    
	PrintLog "txtSpread2 = " & Request("txtSpread2")
	
	arrRowVal = Split(Request("txtSpread2"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        If arrColVal(C_SEQ_NO) <> "999999" Then
				Select Case arrColVal(0)
				     Case "C"
				            Call SubBizSaveMultiCreate2(arrColVal)                            '☜: Create
				     Case "U"
				            Call SubBizSaveMultiUpdate2(arrColVal)                            '☜: Update
				     Case "D"
				            Call SubBizSaveMultiDelete2(arrColVal)                            '☜: Delete
				End Select
        End if
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
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
   
            
	
	lgStrSQL = "INSERT INTO TB_8_3_A WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , W101, W101_NM, W_Law, W102, W_Code" & vbCrLf
	lgStrSQL = lgStrSQL & " , C_W103_AMT, C_W103_RATE_VAL, C_W103_RATE, C_W103 " & vbCrLf
	lgStrSQL = lgStrSQL & " , C_W104, C_Limit_RATE, C_Limit_AMT" & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf

	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W101))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W101_Nm))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_Law))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W102))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_CODE))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W103_AMT), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W103_RATE_VAL), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W103_RATE))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W103))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W104), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_Limit_RATE), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_Limit_AMT), "0"),"0","D")     & "," & vbCrLf

	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate = " & lgStrSQL
	
	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

Function RemovePercent(Byval pVal)
	RemovePercent = Replace(pVal, "%", "")
End Function

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i, iStockCnt
	
	If arrColVal(C_SEQ_NO) <> "999999" Then
		lgStrSQL = "INSERT INTO TB_8_3_B WITH (ROWLOCK) (" & vbCrLf
		lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE  "  & vbCrLf 
		lgStrSQL = lgStrSQL & " ,SEQ_NO,  W105, W105_NM, W106, C_W107, C_W108" & vbCrLf
		lgStrSQL = lgStrSQL & " , C_W109, C_W110, C_W111, C_W112, C_W113" & vbCrLf
		lgStrSQL = lgStrSQL & " , C_W114, C_W115, C_W116, C_W117,C_W118 " & vbCrLf
		lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
		lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
		    
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
        lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W105))),"","S")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W105_NM))),"","S")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W106))),"","S")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W107), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W108), "0"),"0","D")     & "," & vbCrLf
		
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W109), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W110), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W111), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W112), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W113), "0"),"0","D")     & "," & vbCrLf
		
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W114), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W115), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W116), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W117), "0"),"0","D")     & "," & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W118), "0"),"0","D")     & "," & vbCrLf
		
        'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
		'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
		lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
		lgStrSQL = lgStrSQL & ")"
		PrintLog "SubBizSaveMultiCreate = " & lgStrSQL
	
		lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
      End If
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status
	
	
	

    lgStrSQL = "UPDATE  TB_8_3_A WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W101	   = " &  FilterVar(Trim(UCase(arrColVal(C_W101))),"''","S")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W101_NM     = " & FilterVar(Trim(UCase(arrColVal(C_W101_Nm))),"''","S")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W_Law     = " &  FilterVar(Trim(UCase(arrColVal(C_Law))),"''","S")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W102     = " &  FilterVar(Trim(UCase(arrColVal(C_W102))),"''","S")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W_Code     = " &  FilterVar(Trim(UCase(arrColVal(C_CODE))),"''","S")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W103_AMT     = " &  FilterVar(UNICDbl(arrColVal(C_W103_AMT), "0"),"0","D")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W103_RATE_VAL     = " &  FilterVar(UNICDbl(arrColVal(C_W103_RATE_VAL), "0"),"0","D")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W103_RATE     = " &  FilterVar(Trim(UCase(arrColVal(C_W103_RATE))),"''","S")   & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W103     = " & FilterVar(Trim(UCase(arrColVal(C_W103))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W104     = " &  FilterVar(UNICDbl(arrColVal(C_W104), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_Limit_RATE     = " &  FilterVar(UNICDbl(arrColVal(C_Limit_RATE), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_Limit_AMT     = " &  FilterVar(UNICDbl(arrColVal(C_Limit_AMT), "0"),"0","D") & "," & vbCrLf
    

    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W_Code = " & FilterVar(Trim(UCase(arrColVal(C_CODE))),"0","D")  & vbCrLf 

	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiUpdate2
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status


    lgStrSQL = "UPDATE  TB_8_3_B WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W105	   = " &  FilterVar(Trim(UCase(arrColVal(C_W105))),"","S")   & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W105_NM     = " & FilterVar(Trim(UCase(arrColVal(C_W105_NM))),"","S")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W106		= " &  FilterVar(Trim(UCase(arrColVal(C_W106))),"","S")   & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W107     = " &   FilterVar(UNICDbl(arrColVal(C_W107), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W108     = " &  FilterVar(UNICDbl(arrColVal(C_W108), "0"),"0","D")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W109     = " & FilterVar(UNICDbl(arrColVal(C_W109), "0"),"0","D")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W110     = " &  FilterVar(UNICDbl(arrColVal(C_W110), "0"),"0","D")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W111     = " &   FilterVar(UNICDbl(arrColVal(C_W111), "0"),"0","D")   & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W112     = " & FilterVar(UNICDbl(arrColVal(C_W112), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W113     = " &  FilterVar(UNICDbl(arrColVal(C_W113), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W114     = " &   FilterVar(UNICDbl(arrColVal(C_W114), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W115     = " &  FilterVar(UNICDbl(arrColVal(C_W115), "0"),"0","D")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W116     = " &  FilterVar(UNICDbl(arrColVal(C_W116), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W117     = " &  FilterVar(UNICDbl(arrColVal(C_W117), "0"),"0","D")  & "," & vbCrLf
    lgStrSQL = lgStrSQL & " C_W118     = " &  FilterVar(UNICDbl(arrColVal(C_W118), "0"),"0","D")  & "," & vbCrLf
    
    

    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","S")  & vbCrLf




	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_8_3_A WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W101 = " & FilterVar(Trim(UCase(arrColVal(W101))),"0","S")  & vbCrLf 

	PrintLog "SubBizSaveMultiDelete = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_8_3_B WITH (ROWLOCK) "	 & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(C_SEQ_NO))),"0","D")  & vbCrLf 
	' 모두다 삭제하는 루틴(디테일)
	PrintLog "SubBizSaveMultiDelete = " & lgStrSQL
	
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
Sub Window_onload()
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
End Sub
</Script>
<%
'   **************************************************************
'	1.4 Transaction 처러 이벤트 
'   **************************************************************

Sub	onTransactionCommit()
	' 트랜잭션 완료후 이벤트 처리 
End Sub

Sub onTransactionAbort()
	' 트랜잭선 실패(에러)후 이벤트 처리 
'PrintForm
'	' 에러 출력 
	'Call SaveErrorLog(Err)	' 에러로그를 남긴 
	
End Sub
%>
