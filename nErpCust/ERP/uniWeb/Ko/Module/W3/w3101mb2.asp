<%@ Language=VBScript%>
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

	Const TYPE_1 = 0
	Const TYPE_2 = 1

	Dim C_SEQ_NO		'	순번 
	Dim C_ACCT			'	계정명 
	Dim C_W14_1			'	총급여액 
	Dim C_W14_2			'	금액 
	dim C_W15_1			'	1년미만근로한  사용인에 대한 급여액 
	dim C_W15_2			'	금액 
	dim C_W16_1			'	1년간 계속근로한 임원사용인에 대한 급여액 
	dim C_W16_2			'	금액 
	dim C_W17_1			'	기말현재전사용인퇴직시퇴직급여추계액 
	dim C_W17_2			'   금액 
	
	lgErrorStatus    = "NO"
    lgOpModeCRUD     = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    'lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    'lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    'lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

    Call SubOpenDB(lgObjConn) 
    	Call CheckVersion(sFISC_YEAR   ,sREP_TYPE)
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
	 C_SEQ_NO	= 1      ' 
	 C_ACCT		= 2
	 C_W14_1	= 3	
	 C_W14_2	= 4
	 C_W15_1	= 5
	 C_W15_2	= 6
	 C_W16_1	= 7
	 C_W16_2	= 8
	 C_W17_1	= 9
	 C_W17_2	= 10
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(3), iIntMaxRows, iLngRow
    Dim iRow, iKey1, iKey2, iKey3,iKey2_1
	Dim arrRow(2), iType, iStrData, iLngCol
	
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = 1 								' 신고구분(확정인것만)

    iKey2_1 = FilterVar(cdbl(sFISC_YEAR)-1,"''", "S")	' 전연도 

    '그리드 부분 조회 
		' 미리 삭제하지 않으면 중복에러 발생 by sbk
		lgstrData = ""
					
		lgStrSQL = "DELETE  TB_32D"
		lgStrSQL = lgStrSQL & " WHERE        "
		lgStrSQL = lgStrSQL & "  co_cd = " & iKey1  
		lgStrSQL = lgStrSQL & "  and fisc_year =" & iKey2
		lgStrSQL = lgStrSQL & "  and rep_type =" & iKey3 

		'---------- Developer Coding part (End  ) ---------------------------------------------------------------
		lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	
			
		iStrData = ""
		'TYPE_2 조회 
	    Call SubMakeSQLStatements("RD",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
		     lgStrPrevKey = ""
		    'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		    'Call SetErrorStatus()
			    
		Else
				
			lgstrData = ""
	 
			Do While Not lgObjRs.EOF
				iStrData = iStrData & " .Row = .Row " & vbCrLf
				iStrData = iStrData & " .Col = " & C_ACCT & " : .text = """ & ConvSPChars("급여") & """" & vbCrLf
				iStrData = iStrData & " .Col = " & C_W14_1 & " : .value = """ & ConvSPChars(lgObjRs("W2")) & """" & vbCrLf			
				iStrData = iStrData & " .Col = " & C_W14_2 & " : .value = """ & ConvSPChars(lgObjRs("W1")) & """" & vbCrLf	
				iStrData = iStrData & " .Col = " & C_W15_1 & " : .value = """ & 0 & """" & vbCrLf	
				iStrData = iStrData & " .Col = " & C_W15_2 & " : .value = """ & 0 & """" & vbCrLf	
				iStrData = iStrData & " .Col = " & C_W16_1 & " : .value = """ & ConvSPChars(lgObjRs("W4")) & """" & vbCrLf	
				iStrData = iStrData & " .Col = " & C_W16_2 & " : .value = """ & ConvSPChars(lgObjRs("W3")) & """" & vbCrLf	
				iStrData = iStrData & " .Col = " & C_W17_1 & " : .value = """ & 0 & """" & vbCrLf	

				iLngRow = iLngRow + 1
				lgObjRs.MoveNext
			Loop 
	
			lgObjRs.Close
			Set lgObjRs = Nothing
				
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " Dim iRow " & vbCr
			Response.Write " With parent.frm1.vspdData	" & vbCr
			Response.Write "	parent.ggoSpread.Source = parent.frm1.vspdData	 " & vbCr
			Response.Write "	If .MaxRows = 0 Then " & vbCrLf
			Response.Write "		iRow = .MaxRows + 1 " & vbCrLf
			Response.Write "	Else " & vbCrLf
			Response.Write "		iRow = .MaxRows " & vbCrLf
			Response.Write "	End If " & vbCrLf
			Response.Write "	Call parent.FncInsertRow(" & iLngRow & ")" & vbCr
			Response.Write "	.Row = iRow " & vbCrlf
			Response.Write iStrData 
		
			Response.Write " Call parent.FncCalSum(" & C_W14_1 & ",iRow)                                 " & vbCr
			Response.Write " Call parent.FncCalSum(" & C_W14_2 & ",iRow)                                 " & vbCr
			Response.Write " Call parent.FncCalSum(" & C_W16_1 & ",iRow)                                 " & vbCr
			Response.Write " Call parent.FncCalSum(" & C_W16_2 & ",iRow)                                 " & vbCr
			Response.Write " Call parent.txtRemark_Change() " & vbCr

			Response.Write " Call parent.SetSpreadColor(-1,-1)                                " & vbCr
	
			Response.Write "	parent.lgBlnFlgChgValue = true	                        " & vbCr
			Response.Write " End With	                        " & vbCr
			Response.Write " </Script>	                        " & vbCr
		
		End If
		
		
		 Call SubMakeSQLStatements("RD2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
		     lgStrPrevKey = ""
		   ' Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		   ' Call SetErrorStatus()
		    
		Else
			iLngCol = lgObjRs.Fields.Count

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                   " & vbCr
	
			iStrData = " .frm1.vspdData.Row = .frm1.vspdData.Maxrows " & vbCrLf
	        iStrData = iStrData & " .frm1.vspdData.Col = " & C_W17_2 & " : .frm1.vspdData.value = """ & ConvSPChars(lgObjRs("W17")) & """" & vbCrLf	
	        Response.Write		 iStrData 
			Response.Write "	.frm1.txtW4.value = """ & lgObjRs("W4") & """" & vbCrLf	' 데이타출력시 이벤트가 발생하지 못하게 한다.
	  
			Response.Write " End With                                           " & vbCr
			Response.Write " </Script>                                          " & vbCr
     
			lgObjRs.Close
			Set lgObjRs = Nothing
		End If		


	'회사계상액 2007.06
	
	 Call SubMakeSQLStatements("RD3",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
		     lgStrPrevKey = ""
		    
		Else
			iLngCol = lgObjRs.Fields.Count

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                   " & vbCr
			Response.Write "	.frm1.txtW12.value = """ & lgObjRs("W4") & """" & vbCrLf	' 
			Response.Write " End With                                           " & vbCr
			Response.Write " </Script>                                          " & vbCr
     
			lgObjRs.Close
			Set lgObjRs = Nothing
		End If	
		
		
    Call SubMakeSQLStatements("RH",iKey1, iKey2_1, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
       ' Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
       ' Call SetErrorStatus()
        
    Else
		iLngCol = lgObjRs.Fields.Count
	
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr
	
		Response.Write "	.frm1.txtW6.value =  " & unicdbl(lgObjRs("W5"),"0")  &" - .frm1.txtW5.value " & vbCrLf			

		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr

		lgObjRs.Close
		Set lgObjRs = Nothing
	End If
	
	
	Call SubMakeSQLStatements("RH3",iKey1, iKey2_1, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
       ' Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
       ' Call SetErrorStatus()
        
    Else
		iLngCol = lgObjRs.Fields.Count
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                   " & vbCr

		Response.Write "	.frm1.txtRemark.value = """ & lgObjRs("WRemark") & """" & vbCrLf		

		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr

		lgObjRs.Close
		Set lgObjRs = Nothing
	End If		

	    
		 Call SubMakeSQLStatements("RH2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
		     lgStrPrevKey = ""
		   ' Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
		   ' Call SetErrorStatus()
		    
		Else
			iLngCol = lgObjRs.Fields.Count

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                   " & vbCr
	       
			Response.Write "	.frm1.txtW15ho.value = """ & lgObjRs("W15HO") & """" & vbCrLf	' 15호에 입력될 값 
			if uniCdbl(lgObjRs("W7") ,"0") < 0 then
			   Response.Write "	.frm1.txtW7.value = """ & 0 & """" & vbCrLf	' 
			else
			   Response.Write "	.frm1.txtW7.value = """ & lgObjRs("W7") & """" & vbCrLf	' 
			end if   
			   
	  
			Response.Write " End With                                           " & vbCr
			Response.Write " </Script>                                          " & vbCr
     
			lgObjRs.Close
			Set lgObjRs = Nothing
		End If		
	
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

Function GetValueText(Byval pData)
	If Instr(1, pData, "-") > 0 Then
		GetValueText = "text"
	Else
		GetValueText = "value"
	End If
End Function
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "RH"
            '전년도 50호 을의 퇴직급여충당금 - (5)기중충당금 환입액  - > 충당금부인누계액 
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " W5  "
            lgStrSQL = lgStrSQL & " FROM TB_50B  WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "     AND W1= '3201'"

				
	
      Case "RD"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "    A.W1 , A.W2 , A.W3,  A.W4 "
            lgStrSQL = lgStrSQL & " FROM TB_WORK_1_1 A WITH (NOLOCK) "
            lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			
			
	   Case "RD2"    ' 퇴직급여충당금 
	  
	
			lgStrSQL =  " select     Isnull(Sum(CREDIT_BASIC_AMT - DEBIT_BASIC_AMT ),0) W4 ,"
			lgStrSQL = lgStrSQL & "  Isnull(Sum((CREDIT_BASIC_AMT + CREDIT_SUM_AMT)  - ( DEBIT_BASIC_AMT - DEBIT_SUM_AMT )),0)  as W17 "
			lgStrSQL = lgStrSQL & "   from TB_WORK_2 A "
            lgStrSQL = lgStrSQL & "     where  ACCT_CD in (SELECT  ACCT_CD  FROM  TB_ACCT_MATCH "
            lgStrSQL = lgStrSQL & "					WHERE MATCH_CD = '02' AND CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "					AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "					AND REP_TYPE = " & pCode3 &")  "
			lgStrSQL = lgStrSQL & "     AND A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
	
		 Case "RD3"    ' 퇴직급여충당금 회계계상액 
			
			lgStrSQL =  " SELECT  CREDIT_SUM_AMT W4 "
			lgStrSQL = lgStrSQL & "   from TB_WORK_2 A "
            lgStrSQL = lgStrSQL & "     where  ACCT_CD in (SELECT  ACCT_CD  FROM  TB_ACCT_MATCH "
            lgStrSQL = lgStrSQL & "					WHERE MATCH_CD = '02' AND CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "					AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "					AND REP_TYPE = " & pCode3 &")  "
			lgStrSQL = lgStrSQL & "     AND A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			
			
	     Case "RH2"
	
			lgStrSQL =  " select     Isnull(Sum(CREDIT_SUM_AMT ),0) W15HO ,  Isnull(SUM(T.w4) - SUM(CREDIT_SUM_AMT ),0) W7"
			lgStrSQL = lgStrSQL & "   from "
			lgStrSQL = lgStrSQL & "     ( (SELECT CREDIT_SUM_AMT , 0 W4 FROM TB_WORK_2 A  "
		    lgStrSQL = lgStrSQL & "     where  ACCT_CD in (SELECT  ACCT_CD  FROM  TB_ACCT_MATCH "
            lgStrSQL = lgStrSQL & "					WHERE MATCH_CD = '03' AND CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "					AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "					AND REP_TYPE = " & pCode3 &")  "
			lgStrSQL = lgStrSQL & "     AND A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	  &")  "
			lgStrSQL = lgStrSQL & "     UNION ALL "
			lgStrSQL = lgStrSQL & "     ( SELECT 0 DEBIT_SUM_AMT , W4 FROM TB_WORK_1_4 A  "
			lgStrSQL = lgStrSQL & "     WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	  &")) T "
			
		   Case "RH3"
			lgStrSQL =  " select    "
			lgStrSQL = lgStrSQL & "  Isnull(Sum((CREDIT_BASIC_AMT + CREDIT_SUM_AMT)  - ( DEBIT_BASIC_AMT - DEBIT_SUM_AMT )),0)  as WRemark "
			lgStrSQL = lgStrSQL & "   from TB_WORK_2 A "
            lgStrSQL = lgStrSQL & "     where  ACCT_CD in (SELECT  ACCT_CD  FROM  TB_ACCT_MATCH "
            lgStrSQL = lgStrSQL & "					WHERE MATCH_CD = '04' AND CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "					AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "					AND REP_TYPE = " & pCode3 &")  "
			lgStrSQL = lgStrSQL & "     AND A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf	
    End Select

	'PrintLog "SubMakeSQLStatements = " & lgStrSQL
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
             'Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>
