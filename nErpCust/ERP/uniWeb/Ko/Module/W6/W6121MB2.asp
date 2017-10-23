<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<%Option Explicit%>
<% session.CodePage=949 %> 
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

	
	Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.
	
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
    	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
    End Select

    Call SubCloseDB(lgObjConn)
    

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
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
	Dim iDx, arrRs(1), iIntMaxRows, iLngRow,iLngCol
    Dim iRow, iKey1, iKey2, iKey3
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 


		
		Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements
 
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
		     lgStrPrevKey = ""
		    Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
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
						                         lgstrData = lgstrData & "  .Col = " &  C_W105    & ": .CellType = 1    " & vbCrLf
												 lgstrData = lgstrData & "  .Col = " &  C_W105_NM & ": .CellType = 1 : .TypeMaxEditLen  = 100 " & vbCrLf
												 
									
                                            end if
								          
									
											 
											
											if  right( ConvSPChars(lgObjRs("W105") ) ,2 ) ="99" then
											    lgstrData = lgstrData & "	.Col = " & C_SEQ_NO          & " : .value   = """ & 999999   & """" & vbCrLf
											    lgstrData = lgstrData & "	.Col = " & C_W105			 & " : .value   = """ & 99 & """" & vbCrLf
											else
											    lgstrData = lgstrData & "	.Col = " & C_SEQ_NO          & " : .value   = """ &  iDx   & """" & vbCrLf
											    lgstrData = lgstrData & "	.Col = " & C_W105			 & " : .text   = """ & ConvSPChars(lgObjRs("W105")) & """" & vbCrLf
											    lgstrData = lgstrData & "   Call parent.vspdData_Change(" & TYPE_2 & "," & C_W105_NM & " ," &iDx & ") " & vbCrLf
											end if    
										
											lgstrData = lgstrData & "	.Col = " & C_W105_NM		 & " : .text   = """ & ConvSPChars(lgObjRs("W105_NM")) & """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W106			 & " : .text   = """ & ConvSPChars(lgObjRs("W106"))  & """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W107			 & " : .value   = """ & ConvSPChars(lgObjRs("W107")) & """" & vbCrLf
											
											lgstrData = lgstrData & "	.Col = " & C_W108			 & " : .value   = """ & ConvSPChars(lgObjRs("W108")) & """" & vbCrLf
											lgstrData = lgstrData & "	.Col = " & C_W109			 & " : .value   = """ & ConvSPChars(lgObjRs("W107")) & """" & vbCrLf
											
											lgstrData = lgstrData & "    Call parent.FncSumSheet(parent.lgvspdData(" & TYPE_2 & ") , "&iDx & ", " & C_W109 & " , " & C_W113 & " , " & true  & " , " &iDx & " ," & C_W114 &" , ""H"")  " & vbCrLf
											if  right( ConvSPChars(lgObjRs("W105") ) ,2 ) <>"99" then
											     
											      lgstrData = lgstrData & "    Call parent.SetColSum3(" & iDx & " )  " & vbCrLf
											      lgstrData = lgstrData & "    Call parent.SetColSum2(" & iDx & " )  " & vbCrLf
											end if   
											lgstrData = lgstrData & "	.Col = " & C_W118 + 1 & " : .value =    """  &  iDx  & """" & vbCrLf
									        
                                            strGubn =  ConvSPChars(lgObjRs("W105"))
								If Err.number <> 0 Then
									PrintLog "iDx=" & iDx
									Exit Sub
								End If
						        iDx = iDx +1    
								lgObjRs.MoveNext
								
						         
							Loop
							
							lgstrData = lgstrData & "	Call parent.CheckReCalc() " & vbCrLf
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
      Case "R2"
			
			
			
			
			lgStrSQL =  " SELECT  SEQ_NO , W105  ,W105_NM, W106 ,  sum(w107) W107 , sum(W108) W108  "
			lgStrSQL = lgStrSQL & " FROM ( "
			lgStrSQL = lgStrSQL & " SELECT "
            lgStrSQL = lgStrSQL & " '' SEQ_NO   , A.W_CODE as W105, A.W101_NM as W105_NM, fisc_END_dt as  W106, A.C_W104  as w107 , 0 AS W108"
			lgStrSQL = lgStrSQL & " FROM TB_8_3_A A WITH (NOLOCK)  , TB_COMPANY_HISTORY B "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	  & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	  & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = B.FISC_YEAR	 " & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = B.REP_TYPE "	& vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.CO_CD = B.CO_CD	 " & vbCrLf
			lgStrSQL = lgStrSQL & "		AND  A.W_CODE <> 30 and  A.C_W104 <> 0"  & vbCrLf
			lgStrSQL = lgStrSQL & "  union all"
			lgStrSQL = lgStrSQL & " SELECT '' SEQ_NO   ,  W105,  W105_NM, W106 as  W106, 0 AS w107,  C_W118 AS W108"
			lgStrSQL = lgStrSQL & " FROM TB_8_3_B A WITH (NOLOCK) , TB_COMPANY_HISTORY B "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR		=   CAST(" & pCode2 &" AS INT) -1		"      & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE		= 1 "						  & vbCrLf
	
			lgStrSQL = lgStrSQL & "		AND B.FISC_YEAR =" & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND B.REP_TYPE = " & pCode3 	& vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.CO_CD = B.CO_CD	 " & vbCrLf
			
			
			lgStrSQL = lgStrSQL & "		AND  A.W105 <> 30 and  A.C_w118 <> 0"  & vbCrLf
			lgStrSQL = lgStrSQL & "  union all" 
			lgStrSQL = lgStrSQL & "  SELECT 9999999 SEQ_NO ,  A.W_CODE +'Sum99'   as W105, '소계' as W105_NM,NULL W106,Sum(A.C_W104)  as w107  ,0 AS W108"
			lgStrSQL = lgStrSQL & "  FROM TB_8_3_A A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & "  WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND  A.W_CODE <> 30  and  A.C_W104 <> 0 "  & vbCrLf
			lgStrSQL = lgStrSQL & "  GROUP BY   A.W_CODE,  A.W101_NM " & vbcrlf
			
			lgStrSQL = lgStrSQL & "  union all"
			lgStrSQL = lgStrSQL & " SELECT  9999999 SEQ_NO    ,  W105  +'Sum99'   as W105,  '소계' as W105_NM,  NULL  W106, 0 AS w107,  Sum(C_W118)  AS W108"
			lgStrSQL = lgStrSQL & " FROM TB_8_3_B A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR		=   CAST(" & pCode2 &" AS INT) -1		"      & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE		= 1 "						  & vbCrLf
		
			lgStrSQL = lgStrSQL & "  GROUP BY   A.W105,  A.W106 " & vbcrlf
			
			lgStrSQL = lgStrSQL & "	 union all"
			lgStrSQL = lgStrSQL & "	 SELECT 9999999 SEQ_NO ,  '99999'  as W105 , '계' as W105_NM,NULL W106 ,Sum(A.C_W104)  as w107 , 0 AS W108"
			lgStrSQL = lgStrSQL & "  FROM TB_8_3_A A WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & "  WHERE A.CO_CD = " & pCode1 	 & vbCrLf 
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND  A.W_CODE <> 30 and  A.C_W104 <> 0 "  & vbCrLf
			lgStrSQL = lgStrSQL & " ) T "
			lgStrSQL = lgStrSQL & " GROUP BY   SEQ_NO , W105  ,W105_NM , W106"
			lgStrSQL = lgStrSQL & " ORDER by  W105 , SEQ_NO , W105_NM , W106 "
			
			
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