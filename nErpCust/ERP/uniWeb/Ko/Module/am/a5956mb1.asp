<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

    On Error Resume Next
    Err.Clear

	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")   
    Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")

	Dim LngMaxRow_i								   										'☜: 최대 업데이트된 갯수 
	Dim lgInsureCD
    Dim Str_Spread
	Dim lgCurrency    
    Dim lgDoc_Cur
    Dim iLoopCount

	Const C_SHEETMAXROWS_D  = 100

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	
    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

	
	'Multi
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UniCLng(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

	lgCurrency		=	Trim(Request("lgCurrency"))
	lgDoc_Cur		=	UCase(Request("txtTradeCur"))
	
	
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	Call SubCreateCommandObject(lgObjComm)
	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                      '☜: Save,Update
			 Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
			
    End Select
			  
	Call SubCloseCommandObject(lgObjComm)
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
    
 

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1
    Dim TEMP_GL_DT, CNT_FROM_DT, CNT_TO_DT, FROM_DT, TO_DT 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	
    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    
    lgInsureCD = iKey1


    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements
	
	IF lgStrPrevKeyIndex = 0 Then
				
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
			
			If lgPrevNext = "" Then
				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
%>
<Script Language=vbscript> 
			
				Parent.Frm1.txtInsuerCd.Value		= ""                   'Set condition area
				Parent.Frm1.txtInsuerNm.Value		= ""
</Script>       
<%
			ElseIf lgPrevNext = "P" Then
			   Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the starting data. 
			   lgPrevNext = ""
			   Call SubBizQuery()
%>
<Script Language=vbscript> 
				Parent.Frm1.UserPrevNext.Value		= "P"
</Script>       
<%          
			ElseIf lgPrevNext = "N" Then
			   Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the ending data.
			   lgPrevNext = ""
			   Call SubBizQuery()
%>
<Script Language=vbscript> 
			Parent.Frm1.UserPrevNext.Value		= "N"
</Script>       
<%          
			End If            
		Else
			
			TEMP_GL_DT	= UNIConvYYYYMMDDToDate(gDateFormat, Trim(lgObjRs(6)), Trim(lgObjRs(7)), Trim(lgObjRs(8)))
			CNT_FROM_DT = UNIConvYYYYMMDDToDate(gDateFormat, Trim(lgObjRs(12)), Trim(lgObjRs(13)), Trim(lgObjRs(14)))
			CNT_TO_DT	= UNIConvYYYYMMDDToDate(gDateFormat, Trim(lgObjRs(15)), Trim(lgObjRs(16)), Trim(lgObjRs(17)))
			FROM_DT		= UNIConvYYYYMMDDToDate(gServerDateFormat, Trim(lgObjRs(22)), Trim(lgObjRs(23)), "01")
			TO_DT		= UNIConvYYYYMMDDToDate(gServerDateFormat, Trim(lgObjRs(24)), Trim(lgObjRs(25)), "01")
		
			lgCurrency  = Trim(lgObjRs("DOC_CUR"))
%>
<Script Language=vbscript>
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	' Set condition area, contents area
	'--------------------------------------------------------------------------------------------------------
		With Parent
				 .Frm1.txtInsuerCd.Value		= "<%=ConvSPChars(Trim(lgObjRs("INSURE_CD")))%>"
				 .Frm1.txtInsuerNm.Value		= "<%=ConvSPChars(Trim(lgObjRs("INSURE_NM")))%>"	
		         .Frm1.txtInsuerCd1.Value		= "<%=ConvSPChars(Trim(lgObjRs("INSURE_CD")))%>"                   'Set condition area
		         .Frm1.txtInsuerNm1.Value		= "<%=ConvSPChars(Trim(lgObjRs("INSURE_NM")))%>"
		         .Frm1.txtInsuerTp.Value		= "<%=ConvSPChars(Trim(lgObjRs("INSURE_TYPE")))%>"                   'Set condtent area
		         .Frm1.txtInsuerTpNm.Value		= "<%=ConvSPChars(Trim(lgObjRs("MINOR_NM")))%>"
		         .Frm1.fpDateTime.Text          = "<%=TEMP_GL_DT%>"
		         .Frm1.txtTradeCur.Value		= "<%=ConvSPChars(Trim(lgObjRs("DOC_CUR")))%>"
		         .Frm1.txtTradeCurNm.Value		= "<%=ConvSPChars(Trim(lgObjRs("CURRENCY_DESC")))%>"                   'Set condition area
		         .Frm1.txtCustomCd.Value		= "<%=ConvSPChars(Trim(lgObjRs("CUST_CD")))%>"
		         .Frm1.txtCustomNm.Value		= "<%=ConvSPChars(Trim(lgObjRs("BP_NM")))%>"                   'Set condtent area
		         .Frm1.txtExRate.text			= "<%=UNINumClientFormat(lgObjRs("XCH_RATE"),    ggExchRate.DecPoint, 0)%>"
		         .Frm1.txtCntAmt.text			= "<%=UNINumClientFormat(lgObjRs("CNT_AMT"),    ggAmtOfMoney.DecPoint, 0)%>"                   'Set condition area
		         .frm1.txtLocCntAmt.text		= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("LOC_CNT_AMT"), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"
		         .Frm1.txtDept1.Value           = "<%=ConvSPChars(Trim(lgObjRs("DEPT_CD1")))%>"
		         .Frm1.txtDept1Nm.Value			= "<%=ConvSPChars(Trim(lgObjRs("DEPT_NM1")))%>"
		         .Frm1.txtAmt.text				= "<%=UNINumClientFormat(lgObjRs("AMT"),    ggAmtOfMoney.DecPoint, 0)%>"                   'Set condition area
		         .Frm1.txtDept2.Value			= "<%=ConvSPChars(Trim(lgObjRs("DEPT_CD2")))%>"
		         .Frm1.txtDept2Nm.Value			= "<%=ConvSPChars(Trim(lgObjRs("DEPT_NM2")))%>"                   'Set condtent area
		         .Frm1.txtInternalCd1.Value		= "<%=ConvSPChars(Trim(lgObjRs("INTERNAL_CD1")))%>"
		         .Frm1.txtInternalCd2.Value		= "<%=ConvSPChars(Trim(lgObjRs("INTERNAL_CD2")))%>"
		         .frm1.txtLocAmt.text			= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("LOC_AMT"), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"
		         .Frm1.cboPrivatePublic.Value   = "<%=ConvSPChars(Trim(lgObjRs("END_YN")))%>"
		         .Frm1.txtCntFrom.Text			= "<%=CNT_FROM_DT%>"
		         .Frm1.txtCntTo.Text			= "<%=CNT_TO_DT%>"                   'Set condition area
		         .Frm1.txtFromDt.Text			= "<%=UNIMonthClientFormat(FROM_DT)%>"
		         .Frm1.txtToDt.Text				= "<%=UNIMonthClientFormat(TO_DT)%>"                   'Set condtent area
		         .Frm1.txtRefNo.Value			= "<%=ConvSPChars(Trim(lgObjRs("REF_NO")))%>"
		         .Frm1.txtTempGlNo.Value        = "<%=ConvSPChars(Trim(lgObjRs("TEMP_GL_NO")))%>"
		         .Frm1.txtGlNo.Value			= "<%=ConvSPChars(Trim(lgObjRs("GL_NO")))%>"
		         .Frm1.txtCostCd.Value			= "<%=ConvSPChars(Trim(lgObjRs("BIZ_AREA_CD")))%>"
				 .Frm1.txtOrgChId.Value			= "<%=ConvSPChars(Trim(lgObjRs("ORG_CHANGE_ID")))%>"
				 .Frm1.txtInsureAcct.Value		= "<%=ConvSPChars(Trim(lgObjRs("ACCT_CD")))%>"
				 .Frm1.txtInsureAcctNm.Value	= "<%=ConvSPChars(Trim(lgObjRs("ACCT_NM")))%>"
		End With          
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
</Script>       
<%     
	   End If

		
		
		Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
	

	END IF  
	
	
End Sub	

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db (A_INSURE_ITEM)
'============================================================================================================
Sub SubBizQueryMulti()

	Dim iDx
	
	
	Call SubMakeSQLStatements("M",lgInsureCD)
	
		
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = 0
    Else
		
        
        lgstrData = ""
        
        If lgStrPrevKeyIndex > 0 Then
			lgObjRs.Move     = C_SHEETMAXROWS_D* lgStrPrevKeyIndex                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		End If

         
        iLoopCount = 0 
        
                
        Do While Not (lgObjRs.EOF Or lgObjRs.BOF)
        
           If  iLoopCount < C_SHEETMAXROWS_D Then
            iLoopCount =  iLoopCount + 1
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SEQ"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_TYPE"))
            lgstrData = lgstrData & Chr(11) & " "
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("PAY_AMT"),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("PAY_LOC_AMT"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NOTE_NO"))
            lgstrData = lgstrData & Chr(11) & " "
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BANK_ACCT_NO"))
			lgstrData = lgstrData & Chr(11) & " "
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BANK_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BANK_NM"))
 			       
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iLoopCount
            lgstrData = lgstrData & Chr(11) & Chr(12)
			
			Else
				lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
				Exit Do
			End If
        		
		    lgObjRs.MoveNext
        Loop	
    End If     
   
    If  iLoopCount < C_SHEETMAXROWS_D Then             

                                   '☜: Check if next data exists
        lgStrPrevKeyIndex = 0                                                  '☜: 다음 데이타 없다.

    End If
  	
  	
  	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)  
End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = UniCInt(Request("txtFlgMode"),0)                                       '☜: Read Operayion Mode (CREATE, UPDATE)

	Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                         '☜ : Create
			  Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE
              Call SubBizSaveSingleUpdate()
    End Select

End Sub	

'============================================================================================================
' Name : SubBizBatch
' Desc : Date data 
'============================================================================================================
Sub SubBizBatch(iOption)

    Dim Maxprice, LngRecs
	Dim IntRetCD
    Dim gl_date
    Dim insure_cd
	DIM strMsg_cd
   
   On Error Resume Next                                                             '☜: Protect system from crashing
   Err.Clear                                                                        '☜: Clear Error status

     Call SubCreateCommandObject(lgObjComm)
     
	If UCase(iOption) <>  "D" Then
		gl_date = Trim(lgKeyStream(4))
		insure_cd = UCase(Trim(Request("txtInsuerCd1")))
	ELSE
		gl_date = Trim(lgKeyStream(3))
		insure_cd = UCase(lgKeyStream(0))
	END If	  
'	Call ServerMesgBox("SP전||" & iOption &  "||" & insure_cd &"||" &  gl_date &"||" & gUsrId  , vbInformation, I_MKSCRIPT)
	With lgObjComm
	    .CommandText = "A_USP_A5955BA1_INSURE_POSTING"
	    .CommandType = adCmdStoredProc
			
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@iOption"   ,adVarWChar,adParamInput, 1, iOption)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@insure_cd"   ,adVarWChar,adParamInput, 20, insure_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@temp_gl_date"   ,adVarWChar,adParamInput, 8, gl_date)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id",adVarWChar,adParamInput,13, gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",adVarWChar,adParamOutput,6)
		 
		lgObjComm.Execute ,, adExecuteNoRecords
		    
	End With
'	Call ServerMesgBox("SP후||" & iOption &  "||" & insure_cd &"||" &  gl_date &"||" & gUsrId  , vbInformation, I_MKSCRIPT)
	If Err.number <> 0 Then
		ObjectContext.SetAbort
		Err.Clear      
	End If	
	If  Err.number = 0 Then
	   IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
       if  IntRetCD <> 1 then
	        strMsg_cd = lgObjComm.Parameters("@msg_cd").Value            
            Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
			ObjectContext.SetAbort		'//추가 
            Response.End
            Exit Sub
       end if
    Else
	    lgErrorStatus     = "YES"     
        ObjectContext.SetAbort
        Call SubHandleError("Batch", lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
        
    End if
    Call SubCloseCommandObject(lgObjComm)
    
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
   
    lgStrSQL =  "DELETE  A_INSURE_ITEM"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " INSURE_CD   = " & FilterVar(lgKeyStream(0), "''", "S")
    
    
    lgStrSQL = lgStrSQL & " DELETE  A_INSURE"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " INSURE_CD   = " & FilterVar(lgKeyStream(0), "''", "S")
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
	'Response.Write lgStrSQL & vbcrlf
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)
	

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

	Dim CNT_FROM_DT, CNT_TO_DT, FROM_DT, TO_DT, TEMP_GL_DT, pCONST
	Dim strYear,strMonth,strDay
	Dim loc_cnt_amt, loc_amt
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Call SubMakeSQLStatements("Q",FilterVar(UCase(Request("txtInsuerCd1")), "''", "S"))                                 '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
           Call DisplayMsgBox("800446", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
           Response.End 
        Call SetErrorStatus()
    End If


	loc_cnt_amt =  GetLocAmt(lgDoc_Cur,gCurrency,lgKeyStream(4),UNIConvNum(Request("txtCntAmt"),0),"AMT") '//sp타는 함수 적용시켜야함 
	loc_amt =  GetLocAmt(lgDoc_Cur,gCurrency,lgKeyStream(4),UNIConvNum(Request("txtAmt"),0),"AMT") '//sp타는 함수 적용시켜야함 
	pCONST = 0
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = lgStrSQL & "  INSERT INTO A_INSURE( "
    lgStrSQL = lgStrSQL & " INSURE_CD     ," 
    lgStrSQL = lgStrSQL & " INSURE_NM     ," 
    lgStrSQL = lgStrSQL & " ACCT_CD     ,"
    lgStrSQL = lgStrSQL & " INSURE_TYPE     ," 
    lgStrSQL = lgStrSQL & " CUST_CD     ," 
    lgStrSQL = lgStrSQL & " DOC_CUR     ," 
    lgStrSQL = lgStrSQL & " XCH_RATE     ," 
    lgStrSQL = lgStrSQL & " CNT_FROM_DT     ,"
    lgStrSQL = lgStrSQL & " CNT_TO_DT     ," 
    lgStrSQL = lgStrSQL & " CNT_AMT     ," 
    lgStrSQL = lgStrSQL & " LOC_CNT_AMT     ,"  
    lgStrSQL = lgStrSQL & " AMT     ,"
    lgStrSQL = lgStrSQL & " LOC_AMT     ," 
    lgStrSQL = lgStrSQL & " FROM_DT     ," 
    lgStrSQL = lgStrSQL & " TO_DT     ," 
    lgStrSQL = lgStrSQL & " BIZ_AREA_CD     ," 
    lgStrSQL = lgStrSQL & " ORG_CHANGE_ID     ," 
    lgStrSQL = lgStrSQL & " DEPT_CD1     ," 
    lgStrSQL = lgStrSQL & " DEPT_CD2    ," 
    lgStrSQL = lgStrSQL & " INTERNAL_CD1     ," 
    lgStrSQL = lgStrSQL & " INTERNAL_CD2    ," 
    lgStrSQL = lgStrSQL & " END_YN         ," 
    lgStrSQL = lgStrSQL & " REF_NO," 
    lgStrSQL = lgStrSQL & " TEMP_GL_DT     ," 
    lgStrSQL = lgStrSQL & " TEMP_GL_NO ,"
    lgStrSQL = lgStrSQL & " INSRT_USER_ID     ," 
    lgStrSQL = lgStrSQL & " INSRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID     ,"  
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtInsuerCd1")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtInsuerNm1")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtInsureAcct"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtInsuerTp"), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtCustomCd")), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtTradeCur")), "''", "S")             & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtExRate"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S")              & ","
   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtCntAmt"),0) & "," 
    lgStrSQL = lgStrSQL & loc_cnt_amt & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtAmt"),0) & ","
    lgStrSQL = lgStrSQL & loc_amt & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(2), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtCostCd")), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtOrgChId")), "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtDept1")), "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtDept2")), "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtInternalCd1")), "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtInternalCd2")), "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("cboPrivatePublic")), "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtRefNo"), "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(4), "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(last_auto_no, "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                             & "," 
    lgStrSQL = lgStrSQL & "getdate()"      & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                             & "," 
    lgStrSQL = lgStrSQL & "getdate()"
    lgStrSQL = lgStrSQL & ")"
    

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    Dim loc_cnt_amt, loc_amt
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    If Request("txtGlNo") <> "" Then
		lgStrSQL = " UPDATE  A_INSURE "
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " DEPT_CD1   =		" & FilterVar(UCase(Request("txtDept1")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " DEPT_CD2   =		" & FilterVar(UCase(Request("txtDept2")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " END_YN   =			" & FilterVar(UCase(Request("cboPrivatePublic")), "''", "S")		& ","
		lgStrSQL = lgStrSQL & " UPDT_USER_ID   =	" & FilterVar(gUsrId, "''", "S")									& "," 
		lgStrSQL = lgStrSQL & " UPDT_DT   =			getdate() " 
		lgStrSQL = lgStrSQL & " WHERE				"
		lgStrSQL = lgStrSQL & " INSURE_CD   =		" & FilterVar(UCase(Request("txtInsuerCd1")), "''", "S")
		

    Else 
			
		loc_cnt_amt =  GetLocAmt(UCase(Request("txtTradeCur")),gCurrency,lgKeyStream(4),UNIConvNum(Request("txtCntAmt"),0),"AMT") '//sp타는 함수 적용시켜야함 
		loc_amt =  GetLocAmt(UCase(Request("txtTradeCur")),gCurrency,lgKeyStream(4),UNIConvNum(Request("txtAmt"),0),"AMT") '//sp타는 함수 적용시켜야함 

	
		lgStrSQL = lgStrSQL & "      UPDATE  A_INSURE "
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " INSURE_NM   =		" & FilterVar(UCase(Request("txtInsuerNm1")), "''", "S")				& ","
		lgStrSQL = lgStrSQL & " ACCT_CD   =		" & FilterVar(Request("txtInsureAcct"), "''", "S")				& ","
		lgStrSQL = lgStrSQL & " INSURE_TYPE   =		" & FilterVar(UCase(Request("txtInsuerTp")), "''", "S")          & ","
		lgStrSQL = lgStrSQL & " CUST_CD   =			" & FilterVar(UCase(Request("txtCustomCd")), "''", "S")          & ","
		lgStrSQL = lgStrSQL & " DOC_CUR   =			" & FilterVar(UCase(Request("txtTradeCur")), "''", "S")          & ","
		lgStrSQL = lgStrSQL & " XCH_RATE   =		" & UniConvNum(Request("txtExRate"),0)							& ","
		lgStrSQL = lgStrSQL & " CNT_FROM_DT   =		" & FilterVar(lgKeyStream(0), "''", "S")							& ","
		lgStrSQL = lgStrSQL & " CNT_TO_DT   =		" & FilterVar(lgKeyStream(1), "''", "S")							& ","
		lgStrSQL = lgStrSQL & " CNT_AMT   =			" & UniConvNum(Request("txtCntAmt"),0)						& "," 
		lgStrSQL = lgStrSQL & " LOC_CNT_AMT   =		" & loc_cnt_amt						& "," 
		lgStrSQL = lgStrSQL & " AMT			=		" & UniConvNum(Request("txtAmt"),0)						& "," 
		lgStrSQL = lgStrSQL & " LOC_AMT			=	" & loc_amt							& "," 
		lgStrSQL = lgStrSQL & " FROM_DT   =			" & FilterVar(lgKeyStream(2), "''", "S")							& ","
		lgStrSQL = lgStrSQL & " TO_DT   =			" & FilterVar(lgKeyStream(3), "''", "S")							& ","
		lgStrSQL = lgStrSQL & " BIZ_AREA_CD   =		" & FilterVar(UCase(Request("txtCostCd")), "''", "S")            & ","
		lgStrSQL = lgStrSQL & " ORG_CHANGE_ID   =	" & FilterVar(UCase(Request("txtOrgChId")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " DEPT_CD1   =		" & FilterVar(UCase(Request("txtDept1")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " DEPT_CD2   =		" & FilterVar(UCase(Request("txtDept2")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " INTERNAL_CD1   =	" & FilterVar(UCase(Request("txtInternalCd1")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " INTERNAL_CD2   =	" & FilterVar(UCase(Request("txtInternalCd2")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " END_YN   =			" & FilterVar(UCase(Request("cboPrivatePublic")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " REF_NO   =			" & FilterVar(Request("txtRefNo"), "''", "S")					& ","
		lgStrSQL = lgStrSQL & " TEMP_GL_DT     =	" & FilterVar(lgKeyStream(4), "''", "S")							& "," 
		lgStrSQL = lgStrSQL & " UPDT_USER_ID   =	" &	 FilterVar(gUsrId, "''", "S")									& "," 
		lgStrSQL = lgStrSQL & " UPDT_DT   =	getdate()	"   
		lgStrSQL = lgStrSQL & " WHERE INSURE_CD   =		" & FilterVar(UCase(Request("txtInsuerCd1")), "''", "S") 
		
	

    End If
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
	
	
    
End Sub

'///////////////////////////exchange rate and local amount //////////////////////////////////////////////////////////////////////
'============================================================================================================
' Name : GetLocAmt(from_cur,to_cur,apprl_dt,fr_value, flg)
' Desc : 현재의 자국금액 
'============================================================================================================	
Function GetLocAmt(from_cur,to_cur,apprl_dt,fr_value, flg)

    On Error Resume Next    
    Dim IntRetCD
    CONST CALLSPNAME = "usp_c_trans_exchange_rate_and_format"
    Call SubCreateCommandObject(lgObjComm)

    With lgObjComm
        .CommandText = CALLSPNAME			'CALLSPNAME
        .CommandType = adCmdStoredProc
		.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
		.Parameters.Append lgObjComm.CreateParameter("@from_cur",adVarWChar,adParamInput,Len(from_cur), from_cur)
		.Parameters.Append lgObjComm.CreateParameter("@to_cur", adVarWChar,adParamInput,Len(Trim(to_cur)), to_cur)
		.Parameters.Append lgObjComm.CreateParameter("@apprl_dt",adVarWChar,adParamInput,Len(apprl_dt), apprl_dt)
		.Parameters.Append lgObjComm.CreateParameter("@fr_value", adVarWChar,adParamInput,Len(Trim(fr_value)), fr_value)
		.Parameters.Append lgObjComm.CreateParameter("@to_value", adVarWChar, adParamOutput, 19)
		.Parameters.Append lgObjComm.CreateParameter("@std_rate", adVarWChar, adParamOutput, 19)
		.Parameters.Append lgObjComm.CreateParameter("@msg_cd", adVarWChar, adParamOutput, 6)
		.Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        If  IntRetCD = 1 then
            to_value		= Trim(lgObjComm.Parameters("@to_value").Value)
            std_rate		= Trim(lgObjComm.Parameters("@std_rate").Value)
            strMsgCd		= Trim(lgObjComm.Parameters("@msg_cd").Value)
			If flg = "AMT" Then
				GetLocAmt		= to_value
			Else
				GetLocAmt		= std_rate
			End If	
			If strMsgCd <> "" Then
			     Call DisplayMsgBox(strMsgCd, vbInformation, "", "", I_MKSCRIPT)
			End If
        end if
    Else 
        lgErrorStatus     = "YES"		                                                  '☜: Set error status
		ObjectContext.SetAbort
       Call SubHandleError("GetLocAmt", lgObjComm.ActiveConnection,lgObjRs,Err)
    End if

	Call SubCloseCommandObject(lgObjComm)

    If lgErrorStatus    = "YES" Then
       lgErrorPos = lgErrorPos & CALLSPNAME & gColSep
    End If
End Function




'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub SuBBizSaveMulti(flg)
	Dim arrRowVal																	'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim arrColVal																	'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt	
	Dim LngRow
	Dim loc_pay_amt
	
	lGrpCnt = 0
	lgStrSQL = ""

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	
	If Request("txtSpread") <> "" Then

		
				
		arrRowVal = Split(Request("txtSpread"), gRowSep)

		lgStrSQL = "Declare @seq smallint "
		
 		For LngRow = 1 To lgLngMaxRow
			
			    
			arrColVal = Split(arrRowVal(LngRow-1), gColSep)
			
			If arrColVal(0) = "Sheet1" Then													'⊙: Sheet 1번 

				lGrpCnt = lGrpCnt + 1														'☜: Group Count

				strStatus = arrColVal(1)														'☜: Row 의 상태 
				
				Select Case strStatus	
				
				Case "D"																'☜: Delete
					lgStrSQL = lgStrSQL	& " Delete A_INSURE_ITEM "
					lgStrSQL = lgStrSQL & " Where INSURE_CD = " & FilterVar(UCase(Request("txtInsuerCd1")), "''", "S")
					lgStrSQL = lgStrSQL & " AND SEQ = " & arrColVal(3)
					
				Case "C"													'☜: Create/Update
					
					loc_pay_amt =  GetLocAmt(lgDoc_Cur,gCurrency,lgKeyStream(4),UniConvNum(arrColVal(5),0),"AMT") '//sp타는 함수 적용시켜야함 
					

					
					lgStrSQL = lgStrSQL	& " Select @seq = isnull(max(seq),0) + 1 from A_INSURE_ITEM  where insure_cd = "
					lgStrSQL = lgStrSQL &	FilterVar(UCase(Request("txtInsuerCd1")), "''", "S")  
										
					lgStrSQL = lgStrSQL	& " Insert into A_INSURE_ITEM (INSURE_CD,SEQ,PAY_TYPE,PAY_AMT,PAY_LOC_AMT,NOTE_NO,BANK_NO,BANK_ACCT_NO,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) Values ("
					lgStrSQL = lgStrSQL	&	FilterVar(UCase(Request("txtInsuerCd1")), "''", "S")      & ","			
					lgStrSQL = lgStrSQL	& " @seq"											        & ","			
					lgStrSQL = lgStrSQL	&	FilterVar(UCase(arrColVal(4)), "''", "S")					& ","			
					lgStrSQL = lgStrSQL	&	UniConvNum(arrColVal(5),0)								& ","			
					lgStrSQL = lgStrSQL	&	loc_pay_amt												& ","			
					lgStrSQL = lgStrSQL	&	FilterVar(UCase(arrColVal(7)), "''", "S")					& ","			
					lgStrSQL = lgStrSQL	&	FilterVar(UCase(arrColVal(8)), "''", "S")					& ","			
					lgStrSQL = lgStrSQL	&	FilterVar(UCase(arrColVal(9)), "''", "S")					& ","			
					lgStrSQL = lgStrSQL &	FilterVar(gUsrId, "''", "S")								& "," 
				    lgStrSQL = lgStrSQL &	"getdate()"					& "," 
					lgStrSQL = lgStrSQL &	FilterVar(gUsrId, "''", "S")								& "," 
					lgStrSQL = lgStrSQL &	"getdate()"
					lgStrSQL = lgStrSQL &	")"

				Case "U"
	
					loc_pay_amt =  GetLocAmt(lgDoc_Cur,gCurrency,lgKeyStream(4),UniConvNum(arrColVal(5),0),"AMT") '//sp타는 함수 적용시켜야함 
					
					lgStrSQL = lgStrSQL	& " Update A_INSURE_ITEM "
					lgStrSQL = lgStrSQL	& " Set  PAY_TYPE = " & FilterVar(UCase(arrColVal(4)), "''", "S")	& ","
					lgStrSQL = lgStrSQL	& " PAY_AMT = " & UniConvNum(arrColVal(5),0)					& ","		
					lgStrSQL = lgStrSQL	& " PAY_LOC_AMT = " & loc_pay_amt								& ","		
					lgStrSQL = lgStrSQL	& " NOTE_NO = " & FilterVar(UCase(arrColVal(7)), "''", "S")		& ","		
					lgStrSQL = lgStrSQL	& " BANK_NO = " & FilterVar(UCase(arrColVal(8)), "''", "S")		& ","		
					lgStrSQL = lgStrSQL	& " BANK_ACCT_NO = " & FilterVar(UCase(arrColVal(9)), "''", "S")		& ","		
					lgStrSQL = lgStrSQL	& " UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S")				& ","		
					lgStrSQL = lgStrSQL	& " UPDT_DT = getdate() " 
					lgStrSQL = lgStrSQL & " Where INSURE_CD = " & FilterVar(UCase(Request("txtInsuerCd1")), "''", "S")
					lgStrSQL = lgStrSQL & " AND SEQ = " & ArrColVal(3)
					
					

				End Select
			End If
		
		Next
		
		
	End If
		'Response.Write lgStrSQL & vbcrlf
    If lgStrSQL <> "" Then
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("SM" & flg,lgObjConn,lgObjRs,Err)
	END If



End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
             
	                 Case ""
                           lgStrSQL = "SELECT A.INSURE_CD, A.INSURE_NM, A.INSURE_TYPE, B.MINOR_NM, A.CUST_CD, C.BP_NM, substring(A.TEMP_GL_DT, 1, 4), substring(A.TEMP_GL_DT, 5, 2), substring(A.TEMP_GL_DT, 7, 2), "
                           lgStrSQL = lgStrSQL & " A.DOC_CUR, D.CURRENCY_DESC, A.XCH_RATE, substring(A.CNT_FROM_DT, 1, 4), substring(A.CNT_FROM_DT, 5, 2), substring(A.CNT_FROM_DT, 7, 2), substring(A.CNT_TO_DT, 1, 4), substring(A.CNT_TO_DT, 5, 2), substring(A.CNT_TO_DT, 7, 2), A.CNT_AMT, "
                           lgStrSQL = lgStrSQL & " A.LOC_CNT_AMT, A.AMT, A.LOC_AMT, substring(A.FROM_DT, 1, 4), substring(A.FROM_DT, 5, 2), substring(A.TO_DT, 1, 4), substring(A.TO_DT, 5, 2), "
                           lgStrSQL = lgStrSQL & " A.DEPT_CD1, G.DEPT_NM AS DEPT_NM1, A.DEPT_CD2, H.DEPT_NM AS DEPT_NM2, INTERNAL_CD1,INTERNAL_CD2,"
                           lgStrSQL = lgStrSQL & " A.END_YN, A.REF_NO, A.TEMP_GL_NO, A.GL_NO, A.BIZ_AREA_CD, A.ORG_CHANGE_ID,A.ACCT_CD, J.ACCT_NM "
                           lgStrSQL = lgStrSQL & " From  A_INSURE A, B_MINOR B, B_BIZ_PARTNER C, B_CURRENCY D, "
                           lgStrSQL = lgStrSQL & " B_ACCT_DEPT G, B_ACCT_DEPT H, B_COMPANY I,A_ACCT J" 
                           lgStrSQL = lgStrSQL & " WHERE A.INSURE_TYPE = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("A1030", "''", "S") & "  "
                           lgStrSQL = lgStrSQL & " AND  A.CUST_CD = C.BP_CD "
                           lgStrSQL = lgStrSQL & " AND A.DOC_CUR = D.CURRENCY "
                           lgStrSQL = lgStrSQL & " AND A.DEPT_CD1 = G.DEPT_CD AND A.DEPT_CD2 = H.DEPT_CD "
                           lgStrSQL = lgStrSQL & " AND G.ORG_CHANGE_ID =* I.CUR_ORG_CHANGE_ID AND H.ORG_CHANGE_ID =* I.CUR_ORG_CHANGE_ID AND A.ACCT_CD = J.ACCT_CD"
                           lgStrSQL = lgStrSQL & " AND A.INSURE_CD = " & pCode 
                     Case "P"
						   lgStrSQL = "SELECT TOP 1 A.INSURE_CD, A.INSURE_NM, A.INSURE_TYPE, B.MINOR_NM, A.CUST_CD, C.BP_NM, substring(A.TEMP_GL_DT, 1, 4), substring(A.TEMP_GL_DT, 5, 2), substring(A.TEMP_GL_DT, 7, 2), "
                           lgStrSQL = lgStrSQL & " A.DOC_CUR, D.CURRENCY_DESC, A.XCH_RATE, substring(A.CNT_FROM_DT, 1, 4), substring(A.CNT_FROM_DT, 5, 2), substring(A.CNT_FROM_DT, 7, 2), substring(A.CNT_TO_DT, 1, 4), substring(A.CNT_TO_DT, 5, 2), substring(A.CNT_TO_DT, 7, 2), A.CNT_AMT, "
                           lgStrSQL = lgStrSQL & " A.LOC_CNT_AMT, A.AMT, A.LOC_AMT, substring(A.FROM_DT, 1, 4), substring(A.FROM_DT, 5, 2), substring(A.TO_DT, 1, 4), substring(A.TO_DT, 5, 2),  "
                           lgStrSQL = lgStrSQL & " A.DEPT_CD1, G.DEPT_NM AS DEPT_NM1, A.DEPT_CD2, H.DEPT_NM AS DEPT_NM2, INTERNAL_CD1,INTERNAL_CD2,"
                           lgStrSQL = lgStrSQL & " A.END_YN, A.REF_NO, A.TEMP_GL_NO, A.GL_NO, A.BIZ_AREA_CD, A.ORG_CHANGE_ID,A.ACCT_CD, J.ACCT_NM "
                           lgStrSQL = lgStrSQL & " From  A_INSURE A, B_MINOR B, B_BIZ_PARTNER C, B_CURRENCY D,"
                           lgStrSQL = lgStrSQL & " B_ACCT_DEPT G, B_ACCT_DEPT H, B_COMPANY I, A_ACCT J " 
                           lgStrSQL = lgStrSQL & " WHERE A.INSURE_TYPE = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("A1030", "''", "S") & "  "
                           lgStrSQL = lgStrSQL & " AND  A.CUST_CD = C.BP_CD "
                           lgStrSQL = lgStrSQL & " AND A.DOC_CUR = D.CURRENCY "
                           lgStrSQL = lgStrSQL & " AND A.DEPT_CD1 = G.DEPT_CD AND A.DEPT_CD2 = H.DEPT_CD "
                           lgStrSQL = lgStrSQL & " AND G.ORG_CHANGE_ID = I.CUR_ORG_CHANGE_ID AND H.ORG_CHANGE_ID = I.CUR_ORG_CHANGE_ID AND A.ACCT_CD = J.ACCT_CD "
                           lgStrSQL = lgStrSQL & " AND A.INSURE_CD < " & pCode
                           lgStrSQL = lgStrSQL & " ORDER BY INSURE_CD DESC "
                     Case "N"
						   lgStrSQL = "SELECT TOP 1 A.INSURE_CD, A.INSURE_NM, A.INSURE_TYPE, B.MINOR_NM, A.CUST_CD, C.BP_NM, substring(A.TEMP_GL_DT, 1, 4), substring(A.TEMP_GL_DT, 5, 2), substring(A.TEMP_GL_DT, 7, 2), "
                           lgStrSQL = lgStrSQL & " A.DOC_CUR, D.CURRENCY_DESC, A.XCH_RATE, substring(A.CNT_FROM_DT, 1, 4), substring(A.CNT_FROM_DT, 5, 2), substring(A.CNT_FROM_DT, 7, 2), substring(A.CNT_TO_DT, 1, 4), substring(A.CNT_TO_DT, 5, 2), substring(A.CNT_TO_DT, 7, 2), A.CNT_AMT, "
                           lgStrSQL = lgStrSQL & " A.LOC_CNT_AMT, A.AMT, A.LOC_AMT, substring(A.FROM_DT, 1, 4), substring(A.FROM_DT, 5, 2), substring(A.TO_DT, 1, 4), substring(A.TO_DT, 5, 2),  "
                           lgStrSQL = lgStrSQL & " A.DEPT_CD1, G.DEPT_NM AS DEPT_NM1, A.DEPT_CD2, H.DEPT_NM AS DEPT_NM2, INTERNAL_CD1,INTERNAL_CD2,"
                           lgStrSQL = lgStrSQL & " A.END_YN, A.REF_NO, A.TEMP_GL_NO, A.GL_NO, A.BIZ_AREA_CD, A.ORG_CHANGE_ID, A.ACCT_CD, J.ACCT_NM "
                           lgStrSQL = lgStrSQL & " From  A_INSURE A, B_MINOR B, B_BIZ_PARTNER C, B_CURRENCY D, "
                           lgStrSQL = lgStrSQL & " B_ACCT_DEPT G, B_ACCT_DEPT H , B_COMPANY I,A_ACCT J" 
                           lgStrSQL = lgStrSQL & " WHERE A.INSURE_TYPE = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("A1030", "''", "S") & "  "
                           lgStrSQL = lgStrSQL & " AND  A.CUST_CD = C.BP_CD "
                           lgStrSQL = lgStrSQL & " AND A.DOC_CUR = D.CURRENCY "
                           lgStrSQL = lgStrSQL & " AND A.DEPT_CD1 = G.DEPT_CD AND A.DEPT_CD2 = H.DEPT_CD "
                           lgStrSQL = lgStrSQL & " AND G.ORG_CHANGE_ID = I.CUR_ORG_CHANGE_ID AND H.ORG_CHANGE_ID = I.CUR_ORG_CHANGE_ID,J.ACCT_NM "
                           lgStrSQL = lgStrSQL & " AND A.INSURE_CD > " & pCode
                           lgStrSQL = lgStrSQL & " ORDER BY INSURE_CD ASC "
                           
             End Select 
      Case "M"		' A_INSURE_ITEM 항목들을 조회 
			lgStrSQL = " Select A.SEQ,A.PAY_TYPE, C.MINOR_NM,isnull(A.PAY_AMT,0) as PAY_AMT,isnull(A.PAY_LOC_AMT,0) as PAY_LOC_AMT,isnull(A.NOTE_NO,'') as NOTE_NO,isnull(A.BANK_NO,'') as BANK_NO,isnull(B.BANK_NM,'') AS BANK_NM,isnull(A.BANK_ACCT_NO,'') AS BANK_ACCT_NO "
			lgStrSQL = lgStrSQL & " From A_INSURE_ITEM A,B_BANK B, B_MINOR C" 
			lgStrSQL = lgStrSQL & " Where A.INSURE_CD = " & pCode  & " AND A.BANK_NO *= B.BANK_CD AND A.PAY_TYPE = C.MINOR_CD AND C.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " "  

			'lgStrSQL = " Select SEQ,PAY_TYPE,isnull(A.PAY_AMT,0) as PAY_AMT,isnull(A.PAY_LOC_AMT,0) as PAY_LOC_AMT,isnull(A.NOTE_NO,'') as NOTE_NO,isnull(A.BANK_NO,'') as BANK_NO,isnull(B.BANK_NM,'') AS BANK_NM,isnull(A.BANK_ACCT_NO,'') AS BANK_ACCT_NO "
			'lgStrSQL = lgStrSQL & " From A_INSURE_ITEM A,B_BANK B " 
			'lgStrSQL = lgStrSQL & " Where A.INSURE_CD = " & pCode  & " AND A.BANK_NO *= B.BANK_CD"  
	  Case "Q"
			lgStrSQL = "SELECT TOP 1 INSURE_CD "
            lgStrSQL = lgStrSQL & " From  A_INSURE "
            lgStrSQL = lgStrSQL & " WHERE INSURE_CD = " & pCode		
     End Select
    
    'Response.Write lgStrSQL & vbcrlf
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
    ObjectContext.SetAbort
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"  
    ObjectContext.SetAbort                                                       '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case left(pOpCode,2)
        Case "SC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    Else
						Call SuBBizSaveMulti("C")
						Call SubBizBatch("C")
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
					Else
						If Trim(lgKeyStream(1)) <> "" or  Trim(lgKeyStream(2)) <> "" Then
							Call SubBizBatch("D")				
						End If				
                    End If
                End If
        Case "SR"
        Case "SM"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "SU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    Else
						Call SuBBizSaveMulti("U")
						Call SubBizBatch("U")
                    End If
                 End If
         Case "SI"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
                 
         Case Else
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
               With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>" ,"F"                         '☜ : Display data
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
				Call .ReFormatSpreadCellByCellByCurrency2(.Frm1.vspdData,"<%=lgLngMaxRow +1%>" , "<%= lgLngMaxRow + iLoopCount%>" ,.frm1.txtTradeCur.value ,.C_Amt ,   "A" ,"I","X","X")

				.DBQueryOk
 	         End with
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
