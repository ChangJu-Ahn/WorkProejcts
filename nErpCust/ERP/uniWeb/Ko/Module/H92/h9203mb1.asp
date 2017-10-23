<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Const C_SHEETMAXROWS_D = 100
	Dim lgStrPrevKey,lgStrPrevKey1
    Dim lgSvrDateTime
    Dim lgTaxFlag, lgLngMaxRow1
    Dim lgSpreadFlg
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
    
    lgSvrDateTime = GetSvrDateTime

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	lgTaxFlag		  = Request("txtTaxFlag")
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgLngMaxRow1      = Request("txtMaxRows1")                                       '☜: Read Operation Mode (CRUD)
    lgSpreadFlg       = Request("lgSpreadFlg")
	if lgSpreadFlg = "1" then
		lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	else		
		lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	end if
    
    Call SubCreateCommandObject(lgObjComm)
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
			 Call SubBizSaveSingleUpdate        
             If lgTaxFlag = "Y" Then ' or lgTaxFlag = "N" Then
				Call SubCreateCommandObject(lgObjComm)
				Call SubBatch()
				Call SubCloseCommandObject(lgObjComm)
			 End If      
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iYymm, iYearType, iEmpNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iYymm     = FilterVar(lgKeyStream(0), "''", "S")
    iYearType = FilterVar(lgKeyStream(1), "''", "S")
    iEmpNo    = FilterVar(lgKeyStream(2), "''", "S")
    
    Call SubMakeSQLStatements("SR",iYymm,iYearType,iEmpNo)                               '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
	   iCnt = iCnt + 1
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the starting data. 
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the ending data.
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
%>
<Script Language=vbscript>
      With Parent.Frm1
             .txtDutyYy.value		= "<%=ConvSPChars(lgObjRs("DUTY_YY"))%>"            
             .txtDutyMm.value		= "<%=ConvSPChars(lgObjRs("DUTY_MM"))%>"            
             .txtDutyDd.value		= "<%=ConvSPChars(lgObjRs("DUTY_DD"))%>"        
             .txtBasAmt.text		= "<%=UNINumClientFormat(lgObjRs("BAS_AMT"), ggAmtOfMoney.DecPoint,0)%>"         
             .txtYearSave.text		= "<%=UNINumClientFormat(lgObjRs("YEAR_SAVE"), ggQty.DecPoint,0)%>"     
             .txtYearPart.text		= "<%=UNINumClientFormat(lgObjRs("YEAR_PART"), ggQty.DecPoint,0)%>"     
             .txtYearBonus.text		= "<%=UNINumClientFormat(lgObjRs("YEAR_BONUS"), ggQty.DecPoint,0)%>"   
             .txtYearSaveTot.text	= "<%=UNINumClientFormat(lgObjRs("YEAR_SAVE_TOT"), ggQty.DecPoint,0)%>"	'2006.04.20
             .txtMaxYearSave.text	= "<%=UNINumClientFormat(lgObjRs("max_year_cnt"), ggQty.DecPoint,0)%>"	'2006.04.20
             .txtYearUse.text		= "<%=UNINumClientFormat(lgObjRs("YEAR_USE"), ggQty.DecPoint,0)%>"     
             .txtYearCnt.text		= "<%=UNINumClientFormat(lgObjRs("YEAR_CNT"), ggQty.DecPoint,0)%>"     
             .txtMonthSave.text		= "<%=UNINumClientFormat(lgObjRs("MONTH_SAVE"), ggQty.DecPoint,0)%>"     
             .txtMonthUse.text		= "<%=UNINumClientFormat(lgObjRs("MONTH_USE"), ggQty.DecPoint,0)%>"     
             .txtMonthDutyCnt.text	= "<%=UNINumClientFormat(lgObjRs("MONTH_DUTY_CNT"), ggQty.DecPoint,0)%>"     
             .txtMonthCnt.text		= "<%=UNINumClientFormat(lgObjRs("MONTH_CNT"), ggQty.DecPoint,0)%>"     
             .txtMonthAmt.text		= "<%=UNINumClientFormat(lgObjRs("MONTH_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
             .txtYearAmt.text		= "<%=UNINumClientFormat(lgObjRs("YEAR_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
             .txtIncomeTaxAmt.text	= "<%=UNINumClientFormat(lgObjRs("INCOME_TAX"), ggAmtOfMoney.DecPoint,0)%>"     
             .txtResTaxAmt.text		= "<%=UNINumClientFormat(lgObjRs("RES_TAX"), ggAmtOfMoney.DecPoint,0)%>"     
             .txtEmpInsurAmt.text	= "<%=UNINumClientFormat(lgObjRs("EMP_INSUR_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
             .txtTotAmt.text		= "<%=UNINumClientFormat(lgObjRs("TOT_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
             .txtRealProvAmt.text	= "<%=UNINumClientFormat(lgObjRs("REAL_PROV_AMT"), ggAmtOfMoney.DecPoint,0)%>"     
      End With   
         
</Script>       

<%     
		Call SubCloseRs(lgObjRs)
		if lgSpreadFlg = "1" then
			If lgErrorStatus <> "YES" Then
'				txtKey = strEmp_no
				Call SubBizQueryMulti1(iYymm,iYearType,iEmpNo)
			End If
		else
		    If lgErrorStatus <> "YES" Then
				'txtKey = strEmp_no
				Call SubBizQueryMulti2(iYymm,iYearType,iEmpNo)
			End If
		end if
    End If
End Sub	

'============================================================================================================
' Name : SubBizQueryMulti1
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti1(pKey1,pKey2,pKey3) 
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MT",pKey1,pKey2,pKey3)                                    '☆: Make sql statements
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""    
		Exit Sub
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		lgstrData1 = ""
        
		iDx = 1
        
		Do While Not lgObjRs.EOF
             			                    
            lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("ALLOW_NM"))
            lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("ALLOW"))
			lgstrData1 = lgstrData1 & Chr(11) & lgLngMaxRow + iDx
			lgstrData1 = lgstrData1 & Chr(11) & Chr(12)

			lgObjRs.MoveNext
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If 
		Loop 

	End If
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If   

    Call SubHandleError("MR",lgObjRs,Err)

    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti2(pKey1,pKey2,pKey3) 
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR",pKey1,pKey2,pKey3)                                 '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey1 = ""    
        Exit Sub
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)    
        lgstrData = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF
                        
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_CD")  )
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_TIME_NM"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("CNT"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey1 = lgStrPrevKey1 + 1
               Exit Do
            End If 
        Loop 
    End If

    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey1 = ""
    End If   

    Call SubHandleError("MR",lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    
  
'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
 '   Err.Clear                                                                        '☜: Clear Error status

   Dim stryear_type
    stryear_type = Request("cboYearType")
   'call svrmsgbox(stryear_type, vbinformation,i_mkscript)

    IF  stryear_type = "2" THEN   '월차 
		lgStrSQL = "UPDATE  HFB020T"
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " MONTH_SAVE = " & UNIConvNum(Request("txtMonthSave"),0) & "," 
		lgStrSQL = lgStrSQL & " MONTH_USE = " & UNIConvNum(Request("txtMonthUse"),0) & ","
		lgStrSQL = lgStrSQL & " MONTH_CNT = " & UNIConvNum(Request("txtMonthCnt"),0) & ","
		lgStrSQL = lgStrSQL & " MONTH_AMT = " & UNIConvNum(Request("txtMonthAmt"),0) & ","
   	    IF lgTaxFlag = "N" THEN 
			lgStrSQL = lgStrSQL & " INCOME_TAX = " & UNIConvNum(Request("txtIncomeTaxAmt"),0) & ","
			lgStrSQL = lgStrSQL & " RES_TAX = " & UNIConvNum(Request("txtResTaxAmt"),0) & ","
			lgStrSQL = lgStrSQL & " EMP_INSUR_AMT = " & UNIConvNum(Request("txtEmpInsurAmt"),0) & ","
		END IF
		lgStrSQL = lgStrSQL & " TAX_AMT = " & UNIConvNum(Request("txtTotAmt"),0) & ","		
    	lgStrSQL = lgStrSQL & " TOT_AMT = " & UNIConvNum(Request("txtTotAmt"),0) & ","
		lgStrSQL = lgStrSQL & " REAL_PROV_AMT = " & UNIConvNum(Request("txtRealProvAmt"),0)               
		lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(lgKeyStream(2), "''", "S")
		lgStrSQL = lgStrSQL & " and YEAR_TYPE = " & FilterVar(lgKeyStream(1), "''", "S")
		lgStrSQL = lgStrSQL & " and YEAR_YYMM = " & FilterVar(lgKeyStream(0), "''", "S")  		
	ELSEIF stryear_type = "1" THEN   '연차 
	    lgStrSQL = "UPDATE  HFB020T"
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " YEAR_SAVE = " & UNIConvNum(Request("txtYearSave"),0) & ","
		lgStrSQL = lgStrSQL & " YEAR_SAVE_TOT = " & UNIConvNum(Request("txtYearSaveTot"),0) & ","
		lgStrSQL = lgStrSQL & " YEAR_PART = " & UNIConvNum(Request("txtYearPart"),0) & ","
		lgStrSQL = lgStrSQL & " YEAR_BONUS = " & UNIConvNum(Request("txtYearBonus"),0) & ","
		lgStrSQL = lgStrSQL & " YEAR_USE = " & UNIConvNum(Request("txtYearUse"),0) & ","
		lgStrSQL = lgStrSQL & " YEAR_CNT = " & UNIConvNum(Request("txtYearCnt"),0) & ","
		lgStrSQL = lgStrSQL & " YEAR_AMT = " & UNIConvNum(Request("txtYearAmt"),0) & ","
   	    IF lgTaxFlag = "N" THEN 
			lgStrSQL = lgStrSQL & " INCOME_TAX = " & UNIConvNum(Request("txtIncomeTaxAmt"),0) & ","
			lgStrSQL = lgStrSQL & " RES_TAX = " & UNIConvNum(Request("txtResTaxAmt"),0) & ","
			lgStrSQL = lgStrSQL & " EMP_INSUR_AMT = " & UNIConvNum(Request("txtEmpInsurAmt"),0) & ","
		END IF
		lgStrSQL = lgStrSQL & " TAX_AMT = " & UNIConvNum(Request("txtTotAmt"),0) & ","			
		lgStrSQL = lgStrSQL & " TOT_AMT = " & UNIConvNum(Request("txtTotAmt"),0) & ","
		lgStrSQL = lgStrSQL & " REAL_PROV_AMT = " & UNIConvNum(Request("txtRealProvAmt"),0)               
		lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(lgKeyStream(2), "''", "S")
		lgStrSQL = lgStrSQL & " and YEAR_TYPE = " & FilterVar(lgKeyStream(1), "''", "S")    
		lgStrSQL = lgStrSQL & " and YEAR_YYMM = " & FilterVar(lgKeyStream(0), "''", "S")  
	END IF 


    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBatch
'============================================================================================================
Sub SubBatch()		
    Dim stryear_yymm
    Dim stryear_yymm_dt
    Dim stryear_type
    Dim strdilig_cd
    Dim strTaxFlag
    Dim strPay_cd
    Dim strallow_cd
    Dim strRetire_flag, strRetire_stdt, strRetire_enddt
    Dim strEmp_no
    Dim strMsg_cd, strMsg_text ,IntRetCD
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    stryear_yymm = lgKeyStream(0)
    stryear_type = Request("cboYearType")

    If  stryear_type  = "1" Then
		Call CommonQueryRs(" ALLOW_CD "," HDA150T "," 1=1 "  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    '연차 
		strallow_cd = Replace(lgF0, Chr(11), "")
    Else
		Call CommonQueryRs(" ALLOW_CD "," HDA140T ", " 1=1  " , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    '월차 
		strallow_cd = Replace(lgF0, Chr(11), "")
    End If		

    strEmp_no = Request("txtEmp_no")
    strTaxFlag = lgTaxFlag
 

    With lgObjComm
        IF  stryear_type  = "1" THEN  '연차이면 
        	IF  strallow_cd <> "" THEN
        		IF  strTaxFlag = "Y" or strTaxFlag = "N" THEN
        		    .CommandText = "usp_hfb030b2"
        		    .CommandType = adCmdStoredProc
                END IF
        	END IF

        ElseIF  stryear_type  = "2" THEN  '월차이면 
        	IF  strallow_cd <> "" THEN
        		IF  strTaxFlag = "Y" or strTaxFlag = "N" THEN
       		        .CommandText = "usp_hfb030b2"
        		    .CommandType = adCmdStoredProc
        		END IF
        	END IF
        ELSE
            IntRetCD = -1
            Exit Sub
        END IF

'Response.Write  "stryear_yymm:" & stryear_yymm
'Response.Write  "stryear_type:" & stryear_type
'Response.Write  "strallow_cd:" & strallow_cd
'Response.Write  "strEmp_no:" & strEmp_no 
'Response.Write  "strTaxFlag:" & strTaxFlag
'Response.End

'call svrmsgbox(strtax_calc  &"/"& strTaxFlag &"/"&  stryear_yymm &"/"& stryear_type & "/"& strallow_cd &"/"& strEmp_no , vbinformation,i_mkscript) 

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adVarXChar,adParamInput, 13 , gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@year_yymm"  ,adVarXChar,adParamInput, 6 ,   stryear_yymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@year_type"  ,adVarXChar,adParamInput, 1 ,   stryear_type)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@allow_cd"   ,adVarXChar,adParamInput, 3 ,    strallow_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",adVarXChar,adParamInput,13 ,      strEmp_no)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@tax_flag"   ,adVarXChar,adParamInput, 1 ,   strTaxFlag)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamoutput, 6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adVarXChar,adParamOutput,60)

        lgObjComm.Execute ,, adExecuteNoRecords 
    End With

    If  Err.number = 0 Then
      IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value  
           
        if  IntRetCD < 0 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value

           Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
             
            IntRetCD = -1
            Exit Sub
        else
            IntRetCD = 1
        end if
    Else        
     
        call svrmsgbox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End if
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode1,pCode2,pCode3)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
                        Select Case  lgPrevNext 
                             Case ""
                                   lgStrSQL =     "Select  a.DUTY_YY, a.DUTY_MM, a.DUTY_DD, a.BAS_AMT, " 
                                   lgStrSQL = lgStrSQL & " a.YEAR_SAVE, a.YEAR_PART, a.YEAR_BONUS, a.YEAR_USE, a.YEAR_CNT, " 
                                   lgStrSQL = lgStrSQL & " a.MONTH_SAVE, a.MONTH_USE, a.MONTH_CNT, a.MONTH_AMT, a.MONTH_DUTY_CNT, " 
                                   lgStrSQL = lgStrSQL & " a.YEAR_AMT, a.INCOME_TAX, a.RES_TAX, a.EMP_INSUR_AMT, " 
                                   lgStrSQL = lgStrSQL & " a.TOT_AMT, a.REAL_PROV_AMT,a.YEAR_SAVE_TOT, b.MAX_YEAR_CNT " 
                                   lgStrSQL = lgStrSQL & " From  HFB020T a, HDA150T b"
                                   lgStrSQL = lgStrSQL & " WHERE YEAR_YYMM = " & pCode1
                                   lgStrSQL = lgStrSQL & "   AND YEAR_TYPE = " & pCode2
                                   lgStrSQL = lgStrSQL & "   AND EMP_NO    = " & pCode3 
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.End
                        End Select
           End Select             
        Case "M" 

           Select Case Mid(pDataType,2,1)
               Case "T"
					   iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1               
					   
                        lgStrSQL = "Select TOP " & iSelCount  & " a.allow_cd, case when allow_nm is null then '합계' else ALLOW_NM end ALLOW_NM, sum(ALLOW) allow" 
                        lgStrSQL = lgStrSQL & " From  HFB030T A, HDA010T B "
                        lgStrSQL = lgStrSQL & " WHERE A.YEAR_YYMM = " & pCode1
                        lgStrSQL = lgStrSQL & "   AND A.YEAR_TYPE = " & pCode2
                        lgStrSQL = lgStrSQL & "   AND A.EMP_NO    = " & pCode3
                        lgStrSQL = lgStrSQL & "   AND A.ALLOW_CD  = B.ALLOW_CD"
                        lgStrSQL = lgStrSQL & "   AND CODE_TYPE = '1' " 	
                        lgStrSQL = lgStrSQL & " GROUP BY a.allow_cd,ALLOW_NM WITH ROLLUP "
                        lgStrSQL = lgStrSQL & " having  not(a.allow_cd is not null and allow_nm is null) " 
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.End
 
                Case "R"
					   iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1               
					   
                        lgStrSQL = "Select TOP " & iSelCount  & " DILIG_CD, dbo.ufn_H_GetCodeName('HCA010T',DILIG_CD,'') DILIG_NM,"
                        lgStrSQL = lgStrSQL & "DAY_TIME,dbo.ufn_GetCodeName('H0086', DAY_TIME) DAY_TIME_NM,CNT "
                        lgStrSQL = lgStrSQL & "FROM HFB040T "
                        lgStrSQL = lgStrSQL & " WHERE YEAR_YYMM = " & pCode1
                        lgStrSQL = lgStrSQL & "   AND YEAR_TYPE = " & pCode2
                        lgStrSQL = lgStrSQL & "   AND EMP_NO    = " & pCode3 
                        

           End Select             
    End Select
 
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
              With Parent
				if .gSpreadFlg = "1" then
					.ggoSpread.Source     = .frm1.vspdData
					.ggoSpread.SSShowData "<%=lgstrData1%>"                               '☜ : Display data
					.lgStrPrevKey    = "<%=lgStrPrevKey%>"
					if .topleftOK then
						.DBQueryOk
					else
						.gSpreadFlg = "2"						
						.DBQuery
					end if
                else
					.ggoSpread.Source     = .frm1.vspdData1
					.ggoSpread.SSShowData "<%=lgstrData%>"          
					.lgStrPrevKey1    = "<%=lgStrPrevKey1%>"
	                .DBQueryOk
                end if
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
