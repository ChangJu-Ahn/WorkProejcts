<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%

    On Error Resume Next
    Err.Clear

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE", "MB")
	Call LoadBNumericformatB("I","*","NOCOOKIE", "MB")

	Dim LngMaxRow_i								   										'☜: 최대 업데이트된 갯수 
    Dim Str_Spread
	Dim lgCurrency 
	Dim lgDoc_Cur
	Dim iLoopCount

    Call HideStatusWnd                                                               '☜: Hide Processing message
	Const C_SHEETMAXROWS_D  = 100
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
	
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Dim strSecuCd
    Dim strSecuNm
    Dim txtSecuCode
    Dim txtTGlNo, txtGlNo
	
	txtTGlNo	= Request("txtTGlNo")
	
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	Call SubCreateCommandObject(lgObjComm)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
			 txtSecuCode	= Request("txtSecuCode")			
			 Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)
			 txtSecuCode	=	lgKeyStream(0)                                                      '☜: Save,Update
             Call SubBizSave()            
        Case CStr(UID_M0003)
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
	Dim lsecurity_cd, lsecurity_nm, ltemp_gl_dt, lsecurity_type
	Dim lsecurity_typeNm, ldoc_cur, lcurrency_desc, lxch_rate
	Dim lbuy_amt, ldept_cd1, ldept_nm1, ldept_cd2
	Dim ldept_nm2, lbiz_area_cd, lorg_change_id, lloc_buy_amt
	Dim linternal_cd1, linternal_cd2, lprice_amt, lcust_cd1
	Dim lbp_nm1, lloc_price_amt, lcust_cd2, lbp_nm2
	Dim lcnt, lend_yn, lpubl_dt, lcalcu_yn
	Dim lint_rate, lcd_mtd, lexpir_dt, lref_no
	Dim lin_dt, ltemp_gl_no, lout_dt, lgl_no
	Dim lacct_cd1, lacct_nm1, lacct_cd2, lacct_nm2

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
	
    'iKey1 = FilterVar(lgKeyStream(0),"''", "S")

    Call SubMakeSQLStatements("MR","X","X",C_EQ)     

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
%>
<Script Language=vbscript> 
       Parent.Frm1.txtSecuCode.Value		= ""                   'Set condition area
       Parent.Frm1.txtSecuNm.Value			= ""
</Script>       
<%          
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the starting data.
          lgPrevNext = ""
          Call SubBizQuery()
%>
<Script Language=vbscript>
		Parent.Frm1.txtPrevNext.Value		= "P"
</Script>
<%
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the ending data.
          lgPrevNext = ""
          Call SubBizQuery()
%>
<Script Language=vbscript>
		Parent.Frm1.txtPrevNext.Value		= "N"
</Script>
<%
       End If
    Else

        lsecurity_cd    = lgObjRs("security_cd")
        lsecurity_nm    = lgObjRs("security_nm")
        ltemp_gl_dt     = lgObjRs("temp_gl_dt")
        lsecurity_type  = lgObjRs("security_type")
        lsecurity_typeNm  = lgObjRs("security_type_nm")
        ldoc_cur        = lgObjRs("doc_cur")
        lcurrency_desc  = lgObjRs("currency_desc")
        lxch_rate       = lgObjRs("xch_rate")
        lbuy_amt        = lgObjRs("buy_amt")
        ldept_cd1       = lgObjRs("dept_cd1")
        ldept_nm1       = lgObjRs("dept_nm1")
        ldept_cd2       = lgObjRs("dept_cd2")
        ldept_nm2       = lgObjRs("dept_nm2")
        lbiz_area_cd    = lgObjRs("biz_area_cd")
        lorg_change_id  = lgObjRs("org_change_id")
        lloc_buy_amt    = lgObjRs("loc_buy_amt")
        linternal_cd1   = lgObjRs("internal_cd1")
        linternal_cd2   = lgObjRs("internal_cd2")
        lprice_amt      = lgObjRs("price_amt")
        lcust_cd1       = lgObjRs("cust_cd1")
        lbp_nm1         = lgObjRs("bp_nm1")
        lloc_price_amt  = lgObjRs("loc_price_amt")
        lcust_cd2       = lgObjRs("cust_cd2")
        lbp_nm2         = lgObjRs("bp_nm2")
        lcnt            = lgObjRs("cnt")
        lend_yn         = lgObjRs("end_yn")
        lpubl_dt        = lgObjRs("publ_dt")
        lcalcu_yn       = lgObjRs("calcu_yn")
        lint_rate       = lgObjRs("int_rate")
        lcd_mtd         = lgObjRs("cd_mtd")
        lexpir_dt       = lgObjRs("expir_dt")
        lref_no         = lgObjRs("ref_no")
        lin_dt          = lgObjRs("in_dt")
        ltemp_gl_no     = lgObjRs("temp_gl_no")
        ltemp_gl_dt     = lgObjRs("temp_gl_dt")
        lout_dt         = lgObjRs("out_dt")
        lgl_no          = lgObjRs("gl_no")
        lacct_cd1		= lgObjRs("acct_cd1")
        lacct_nm1		= lgObjRs("acct_nm1")
		lacct_cd2		= lgObjRs("acct_cd2")
		lacct_nm2		= lgObjRs("acct_nm2")
		lgCurrency  = Trim(lgObjRs("DOC_CUR"))
		
		
		
		
%>
<SCRIPT LANGUAGE=vbscript>
	With Parent
		.frm1.txtSecuCode.value			= "<%=ConvSPChars(lsecurity_cd)%>"
		.frm1.txtSecuNm.value			= "<%=ConvSPChars(lsecurity_nm)%>"
		.frm1.txtSecuCode1.value		= "<%=ConvSPChars(lsecurity_cd)%>"
		.frm1.txtSecuNm1.value			= "<%=ConvSPChars(lsecurity_nm)%>"
        .frm1.txtBillDt.text			= "<%=UNIConvYYYYMMDDToDate(gDateFormat,Mid(ltemp_gl_dt,1,4),Mid(ltemp_gl_dt,5,2),Mid(ltemp_gl_dt,7,2))%>"
        .frm1.txtSecuType.value			= "<%=ConvSPChars(lsecurity_type)%>"
        .frm1.txtSecuTypeNm.value		= "<%=ConvSPChars(lsecurity_typeNm)%>"
        .frm1.txtTradeCur.value			= "<%=ConvSPChars(ldoc_cur)%>"
        .frm1.txtTradeCurNm.value		= "<%=ConvSPChars(lcurrency_desc)%>"
        .frm1.txtXchRate.text			= "<%=UNINumClientFormat(lxch_rate,    ggExchRate.DecPoint, 0)%>"
        .frm1.txtBuyAmt.text			= "<%=UNINumClientFormat(lbuy_amt, ggQty.DecPoint, 0)%>"
        .frm1.txtDept1.value			= "<%=ConvSPChars(ldept_cd1)%>"
        .frm1.txtDept1Nm.value			= "<%=ConvSPChars(ldept_nm1)%>"
        .frm1.txtDept1Area.value		= "<%=ConvSPChars(lbiz_area_cd)%>"
        .frm1.txtDept1OrgId.value		= "<%=ConvSPChars(lorg_change_id)%>"
        .frm1.txtLocBuyAmt.text			= "<%=UNIConvNumDBToCompanyByCurrency(lloc_buy_amt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"
        .frm1.txtDept2.value			= "<%=ConvSPChars(ldept_cd2)%>"
        .frm1.txtDept2Nm.value			= "<%=ConvSPChars(ldept_nm2)%>"
		.frm1.txtInternalCd1.value		= "<%=ConvSPChars(linternal_cd1)%>"
		.frm1.txtInternalCd2.value		= "<%=ConvSPChars(linternal_cd2)%>"
        .frm1.txtPriceAmt.text			= "<%=UNINumClientFormat(lprice_amt, ggQty.DecPoint, 0)%>"
        .frm1.txtCust1.value			= "<%=ConvSPChars(lcust_cd1)%>"
        .frm1.txtCust1Nm.value			= "<%=ConvSPChars(lbp_nm1)%>"
        .frm1.txtLocPriceAmt.text		= "<%=UNIConvNumDBToCompanyByCurrency(lloc_price_amt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"
        .frm1.txtCust2.value			= "<%=ConvSPChars(lcust_cd2)%>"
        .frm1.txtCust2Nm.value			= "<%=ConvSPChars(lbp_nm2)%>"
		.frm1.txtCnt.value				= "<%=UNIConvNumDBToCompanyByCurrency(lcnt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") %>"
        .frm1.selComYn.value			= "<%=ConvSPChars(lend_yn)%>"
        .frm1.txtPubDt.text				= "<%=UNIConvYYYYMMDDToDate(gDateFormat,Mid(lpubl_dt,1,4),Mid(lpubl_dt,5,2),Mid(lpubl_dt,7,2))%>"
        .frm1.selCalYn.value			= "<%=ConvSPChars(lcalcu_yn)%>"
        .frm1.txtCalRate.text			= "<%=UNINumClientFormat(lint_rate, ggQty.DecPoint, 0)%>"
        .frm1.selEndYn.value			= "<%=ConvSPChars(lcd_mtd)%>"
        .frm1.txtExpireDt.text			= "<%=UNIConvYYYYMMDDToDate(gDateFormat,Mid(lexpir_dt,1,4),Mid(lexpir_dt,5,2),Mid(lexpir_dt,7,2))%>"
        .frm1.txtRefNo.value			= "<%=ConvSPChars(lref_no)%>"
        .frm1.txtInDt.text				= "<%=UNIConvYYYYMMDDToDate(gDateFormat,Mid(lin_dt,1,4),Mid(lin_dt,5,2),Mid(lin_dt,7,2))%>"
        .frm1.txtTGlNo.value			= "<%=ConvSPChars(ltemp_gl_no)%>"
        .frm1.txtOutDt.text				= "<%=UNIConvYYYYMMDDToDate(gDateFormat,Mid(lout_dt,1,4),Mid(lout_dt,5,2),Mid(lout_dt,7,2))%>"
        .frm1.txtGlNo.value				= "<%=ConvSPChars(lgl_no)%>"
        .frm1.txtAcct1.value			= "<%=ConvSPChars(lacct_cd1)%>"
        .frm1.txtAcctNm1.value			= "<%=ConvSPChars(lacct_nm1)%>"
        .frm1.txtAcct2.value			= "<%=ConvSPChars(lacct_cd2)%>"
        .frm1.txtAcctNm2.value			= "<%=ConvSPChars(lacct_nm2)%>"
        
	END With
</SCRIPT>
<%
 End If
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db (A_INSURE_ITEM)
'============================================================================================================
Sub SubBizQueryMulti()

	Dim iDx
	
	
	Call SubMakeSQLStatements("MM",FilterVar(txtSecuCode, "''", "S"),"","")
	
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = 0
        
    Else
		
        
        lgstrData = ""
        
        If lgStrPrevKeyIndex > 0 Then
			lgObjRs.Move     = C_SHEETMAXROWS_D * lgStrPrevKeyIndex                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iLoopCount + 1
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
       Case  OPMD_UMODE                                                             '☜ : Update
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
    Dim security_cd
    Dim strMsg_cd

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
	Call SubCreateCommandObject(lgObjComm)

	If iOption <>  "D" Then
		gl_date = Trim(lgKeyStream(25))
		security_cd = UCase(Trim(lgKeyStream(0)))
	ELSE
		gl_date = Trim(lgKeyStream(3))
		security_cd = UCase(Trim(lgKeyStream(0)))
		
	END If	  
	'Call ServerMesgBox("SP전||" & iOption &  "||" & security_cd &"||" &  gl_date &"||" & gUsrId  , vbInformation, I_MKSCRIPT)
	With lgObjComm
	    .CommandText = "A_USP_A5955BA1_SECURITY_POSTING"
	    .CommandType = adCmdStoredProc
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@iOption"   ,adVarWChar,adParamInput, 1, iOption)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@security_cd"   ,adVarWChar,adParamInput, 20, security_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@temp_gl_date"   ,adVarWChar,adParamInput, 8, gl_date)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id",adVarWChar,adParamInput,13, gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",adVarWChar,adParamOutput,6)
		    

	    lgObjComm.Execute ,, adExecuteNoRecords
		    
	End With
	
	'Call ServerMesgBox("SP후||" & iOption &  "||" & security_cd &"||" &  gl_date &"||" & gUsrId  , vbInformation, I_MKSCRIPT)
	If Err.number <> 0 Then
		ObjectContext.SetAbort
		Err.Clear      
	End If	
	If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        if  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value            
            Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
            ObjectContext.SetAbort
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


'///////////////////////////exchange rate and local amount //////////////////////////////////////////////////////////////////////
'============================================================================================================
' Name : GetLocAmt(from_cur,to_cur,apprl_dt,fr_value, flg)
' Desc : 현재의 자국금액 
'============================================================================================================	
Function GetLocAmt(from_cur,to_cur,apprl_dt,fr_value, flg)
	Dim to_value, std_rate, strMsgCd

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
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
	     Call SubHandleError("GetLocAmt", lgObjComm.ActiveConnection,lgObjRs,Err)
    End if

	Call SubCloseCommandObject(lgObjComm)

    If lgErrorStatus    = "YES" Then
       lgErrorPos = lgErrorPos & CALLSPNAME & gColSep
    End If
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    lgStrSQL =  "DELETE  A_SECURITY_ITEM"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " SECURITY_CD   = " & FilterVar(lgKeyStream(0), "''", "S")
   
    lgStrSQL = lgStrSQL & " DELETE  A_SECURITY"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " SECURITY_CD   = " & FilterVar(lgKeyStream(0), "''", "S")
    
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
  '  Response.Write lgstrsql
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
	Dim buy_loc_amt, price_loc_amt
	
	
   On Error Resume Next                                                             '☜: Protect system from crashing
   Err.Clear                                                                         '☜: Clear Error status
	
	Call SubMakeSQLStatements("Q",FilterVar(UCase(Request("txtSecuCode1")), "''", "S"),"X","X")                                 '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
           Call DisplayMsgBox("800446", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
           Response.End 
        Call SetErrorStatus()
    End If

	
	pCONST = 0

	buy_loc_amt =  GetLocAmt(lgDoc_Cur,gCurrency,lgKeyStream(9),UNIConvNum(Request("txtBuyAmt"),0),"AMT") '//sp타는 함수 적용시켜야함 
	price_loc_amt =  GetLocAmt(lgDoc_Cur,gCurrency,lgKeyStream(9),UNIConvNum(Request("txtPriceAmt"),0),"AMT") '//sp타는 함수 적용시켜야함 

	
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
	
    lgStrSQL = lgStrSQL & "  INSERT INTO A_SECURITY( "
    lgStrSQL = lgStrSQL & " SECURITY_CD     ," 
    lgStrSQL = lgStrSQL & " SECURITY_NM     ," 
    lgStrSQL = lgStrSQL & " SECURITY_TYPE     ,"
    lgStrSQL = lgStrSQL & " ACCT_CD1     ,"
    lgStrSQL = lgStrSQL & " ACCT_CD2     ," 
    lgStrSQL = lgStrSQL & " BIZ_AREA_CD     ," 
    lgStrSQL = lgStrSQL & " ORG_CHANGE_ID     ,"
    lgStrSQL = lgStrSQL & " DEPT_CD1     ," 
    lgStrSQL = lgStrSQL & " DEPT_CD2     ,"
	lgStrSQL = lgStrSQL & " INTERNAL_CD1     ," 
    lgStrSQL = lgStrSQL & " INTERNAL_CD2     ,"
    lgStrSQL = lgStrSQL & " PUBL_DT     ,"  
    lgStrSQL = lgStrSQL & " EXPIR_DT     ,"
    lgStrSQL = lgStrSQL & " IN_DT     ," 
    lgStrSQL = lgStrSQL & " CUST_CD1     ," 
    lgStrSQL = lgStrSQL & " CUST_CD2     ," 
    lgStrSQL = lgStrSQL & " DOC_CUR     ," 
    lgStrSQL = lgStrSQL & " XCH_RATE     ," 
    lgStrSQL = lgStrSQL & " BUY_AMT     ," 
    lgStrSQL = lgStrSQL & " LOC_BUY_AMT     ," 
    lgStrSQL = lgStrSQL & " INT_RATE     ," 
    lgStrSQL = lgStrSQL & " CALCU_YN    ," 
    lgStrSQL = lgStrSQL & " CD_MTD         ," 
    lgStrSQL = lgStrSQL & " PRICE_AMT," 
    lgStrSQL = lgStrSQL & " LOC_PRICE_AMT     ," 
    lgStrSQL = lgStrSQL & " CNT ,"
    lgStrSQL = lgStrSQL & " OUT_DT    ," 
    lgStrSQL = lgStrSQL & " END_YN         ," 
    lgStrSQL = lgStrSQL & " REF_NO," 
    lgStrSQL = lgStrSQL & " TEMP_GL_DT," 
    lgStrSQL = lgStrSQL & " INSRT_USER_ID     ," 
    lgStrSQL = lgStrSQL & " INSRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID     ,"  
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtSecuCode1")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtSecuNm1")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtSecuType"), "''", "S")              & ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txtAcct1"), "''", "S")              & ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txtAcct2"), "''", "S")              & ","   
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtDept1Area")), "''", "S") 			& ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtDept1OrgId")), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtDept1")), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtDept2")), "''", "S") 			& ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtInternalCd1")), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtInternalCd2")), "''", "S") 			& ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(7), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(8), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(9), "''", "S")              & ","   
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtCust1")), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtCust2")), "''", "S") 			& ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtTradeCur")), "''", "S") 			& ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtXchRate"),0) & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtBuyAmt"),0) & ","
    lgStrSQL = lgStrSQL & buy_loc_amt & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtCalRate"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("selCalYn")), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("selEndYn")), "''", "S") 			& ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtPriceAmt"),0) & "," 
    lgStrSQL = lgStrSQL & price_loc_amt & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtCnt"),0) & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(22), "''", "S") 			& ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("selComYn")), "''", "S")              & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtRefNo"), "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(25), "''", "S")             & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                             & "," 
    lgStrSQL = lgStrSQL & "getdate()"	     & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                             & "," 
    lgStrSQL = lgStrSQL & "getdate()"
    lgStrSQL = lgStrSQL & ")"
    

'  	Response.Write lgstrsql
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear   
    Dim  buy_loc_amt, price_loc_amt                                                                    '☜: Clear Error status

	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
	
    If Request("txtGlNo") <> "" Then
		lgStrSQL = lgStrSQL & "      UPDATE  A_SECURITY "
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " DEPT_CD1   =		" & FilterVar(UCase(Request("txtDept1")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " DEPT_CD2   =		" & FilterVar(UCase(Request("txtDept2")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " INTERNAL_CD1   =	" & FilterVar(UCase(Request("txtInternalCd1")), "''", "S")           & ","
		lgStrSQL = lgStrSQL & " INTERNAL_CD2   =	" & FilterVar(UCase(Request("txtInternalCd2")), "''", "S") 			& ","
		lgStrSQL = lgStrSQL & " END_YN   =			" & FilterVar(UCase(Request("selComYn")), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " UPDT_USER_ID   =	" & FilterVar(gUsrId, "''", "S")									& "," 
		lgStrSQL = lgStrSQL & " UPDT_DT   =	getdate() " 
		lgStrSQL = lgStrSQL & " WHERE				"
		lgStrSQL = lgStrSQL & " SECURITY_CD   =		" & FilterVar(UCase(Request("txtSecuCode1")), "''", "S")
		
    Else
		buy_loc_amt =  GetLocAmt(lgDoc_Cur,gCurrency,lgKeyStream(9),UNIConvNum(Request("txtBuyAmt"),0),"AMT") '//sp타는 함수 적용시켜야함 
		price_loc_amt =  GetLocAmt(lgDoc_Cur,gCurrency,lgKeyStream(9),UNIConvNum(Request("txtPriceAmt"),0),"AMT") '//sp타는 함수 적용시켜야함 

	
		lgStrSQL = lgStrSQL & "      UPDATE  A_SECURITY "
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " SECURITY_NM   =			" &	 FilterVar(UCase(Request("txtSecuNm1")), "''", "S") & ","
		lgStrSQL = lgStrSQL & " SECURITY_TYPE   =		" & FilterVar(Request("txtSecuType"), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " ACCT_CD1   =		" & FilterVar(Request("txtAcct1"), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " ACCT_CD2   =		" & FilterVar(Request("txtAcct2"), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " BIZ_AREA_CD   =			" & FilterVar(UCase(Request("txtDept1Area")), "''", "S") 			& ","
		lgStrSQL = lgStrSQL & " ORG_CHANGE_ID   =		" & FilterVar(UCase(Request("txtDept1OrgId")), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " DEPT_CD1   =			" & FilterVar(UCase(Request("txtDept1")), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " DEPT_CD2   =			" & FilterVar(UCase(Request("txtDept2")), "''", "S") 			& ","
		lgStrSQL = lgStrSQL & " INTERNAL_CD1   =			" & FilterVar(UCase(Request("txtInternalCd1")), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " INTERNAL_CD2   =			" & FilterVar(UCase(Request("txtInternalCd2")), "''", "S") 			& ","
		lgStrSQL = lgStrSQL & " PUBL_DT   =				" & FilterVar(lgKeyStream(7), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " EXPIR_DT   =			" & FilterVar(lgKeyStream(8), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " IN_DT   =				" & FilterVar(lgKeyStream(9), "''", "S")              & ","   
		lgStrSQL = lgStrSQL & " CUST_CD1   =			" & FilterVar(UCase(Request("txtCust1")), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " CUST_CD2   =			" & FilterVar(UCase(Request("txtCust2")), "''", "S") 			& ","
		lgStrSQL = lgStrSQL & " DOC_CUR   =				" & FilterVar(UCase(Request("txtTradeCur")), "''", "S") 			& ","
		lgStrSQL = lgStrSQL & " XCH_RATE   =			" & UNIConvNum(Request("txtXchRate"),0) & "," 
		lgStrSQL = lgStrSQL & " BUY_AMT   =				" & UNIConvNum(Request("txtBuyAmt"),0) & ","
		lgStrSQL = lgStrSQL & " LOC_BUY_AMT   =			" & buy_loc_amt & ","
		lgStrSQL = lgStrSQL & " INT_RATE   =			" & UNIConvNum(Request("txtCalRate"),0) & ","
		lgStrSQL = lgStrSQL & " CALCU_YN   =			" & FilterVar(UCase(Request("selCalYn")), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " CD_MTD   =				" & FilterVar(UCase(Request("selEndYn")), "''", "S") 			& ","
		lgStrSQL = lgStrSQL & " PRICE_AMT   =			" & UNIConvNum(Request("txtPriceAmt"),0) & "," 
		lgStrSQL = lgStrSQL & " LOC_PRICE_AMT     =		" & price_loc_amt & ","
		lgStrSQL = lgStrSQL & " CNT   =					" & UNIConvNum(Request("txtCnt"),0) & "," 
		lgStrSQL = lgStrSQL & " OUT_DT   =				" & FilterVar(lgKeyStream(22), "''", "S") 			& ","
		lgStrSQL = lgStrSQL & " END_YN   =				" & FilterVar(UCase(Request("selComYn")), "''", "S")              & ","
		lgStrSQL = lgStrSQL & " REF_NO     =			" & FilterVar(Request("txtRefNo"), "''", "S")             & ","
	    lgStrSQL = lgStrSQL & " TEMP_GL_DT =			" &	 FilterVar(lgKeyStream(25), "''", "S")             & ","
		lgStrSQL = lgStrSQL & " UPDT_USER_ID   =		" &	 FilterVar(gUsrId, "''", "S")									& "," 
		lgStrSQL = lgStrSQL & " UPDT_DT   =	getdate()	"
		lgStrSQL = lgStrSQL & " WHERE				"
		lgStrSQL = lgStrSQL & " SECURITY_CD   =			" & FilterVar(UCase(Request("txtSecuCode1")), "''", "S") 
		
    End If
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
	'Response.Write lgstrsql
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
  
End Sub

Sub SuBBizSaveMulti(flg)
	Dim arrRowVal																	'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim arrColVal																	'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt	
	Dim LngRow
	Dim  pay_loc_amt

	
	lGrpCnt = 0
	lgStrSQL = ""

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
	
	
	
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
					lgStrSQL = lgStrSQL	& " Delete A_SECURITY_ITEM "
					lgStrSQL = lgStrSQL & " Where SECURITY_CD = " & FilterVar(UCase(txtSecuCode), "''", "S")
					lgStrSQL = lgStrSQL & " AND SEQ = " & arrColVal(3)
					
				Case "C"													'☜: Create/Update
					pay_loc_amt =  GetLocAmt(lgDoc_Cur,gCurrency,lgKeyStream(9),UniConvNum(arrColVal(5),0),"AMT") '//sp타는 함수 적용시켜야함 
					
					
					lgStrSQL = lgStrSQL	& " Select @seq = isnull(max(seq),0) + 1 from A_SECURITY_ITEM  where security_cd = "
					lgStrSQL = lgStrSQL &	FilterVar(UCase(txtSecuCode), "''", "S")  
										
					lgStrSQL = lgStrSQL	& " Insert into A_SECURITY_ITEM (SECURITY_CD,SEQ,PAY_TYPE,PAY_AMT,PAY_LOC_AMT,NOTE_NO,BANK_NO,BANK_ACCT_NO,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) Values ("
					lgStrSQL = lgStrSQL	&	FilterVar(UCase(txtSecuCode), "''", "S")      & ","			
					lgStrSQL = lgStrSQL	& " @seq"											        & ","			
					lgStrSQL = lgStrSQL	&	FilterVar(UCase(arrColVal(4)), "''", "S")					& ","			
					lgStrSQL = lgStrSQL	&	UniConvNum(arrColVal(5),0)								& ","			
					lgStrSQL = lgStrSQL	&	pay_loc_amt												& ","			
					lgStrSQL = lgStrSQL	&	FilterVar(UCase(arrColVal(7)), "''", "S")					& ","			
					lgStrSQL = lgStrSQL	&	FilterVar(UCase(arrColVal(8)), "''", "S")					& ","			
					lgStrSQL = lgStrSQL	&	FilterVar(UCase(arrColVal(9)), "''", "S")					& ","			
					lgStrSQL = lgStrSQL &	FilterVar(gUsrId, "''", "S")								& "," 
				    lgStrSQL = lgStrSQL &	"getdate()"					& "," 
					lgStrSQL = lgStrSQL &	FilterVar(gUsrId, "''", "S")								& "," 
					lgStrSQL = lgStrSQL &	"getdate()"
					lgStrSQL = lgStrSQL &	")"

				Case "U"
	
					pay_loc_amt =  GetLocAmt(lgDoc_Cur,gCurrency,lgKeyStream(9),UniConvNum(arrColVal(5),0),"AMT") '//sp타는 함수 적용시켜야함 
					
					lgStrSQL = lgStrSQL	& " Update A_SECURITY_ITEM "
					lgStrSQL = lgStrSQL	& " Set  PAY_TYPE = " & FilterVar(UCase(arrColVal(4)), "''", "S")	& ","
					lgStrSQL = lgStrSQL	& " PAY_AMT = " & UniConvNum(arrColVal(5),0)					& ","		
					lgStrSQL = lgStrSQL	& " PAY_LOC_AMT = " & pay_loc_amt								& ","		
					lgStrSQL = lgStrSQL	& " NOTE_NO = " & FilterVar(UCase(arrColVal(7)), "''", "S")		& ","		
					lgStrSQL = lgStrSQL	& " BANK_NO = " & FilterVar(UCase(arrColVal(8)), "''", "S")		& ","		
					lgStrSQL = lgStrSQL	& " BANK_ACCT_NO = " & FilterVar(UCase(arrColVal(9)), "''", "S")		& ","		
					lgStrSQL = lgStrSQL	& " UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S")				& ","		
					lgStrSQL = lgStrSQL	& " UPDT_DT = getdate() " 
					lgStrSQL = lgStrSQL & " Where SECURITY_CD = " & FilterVar(UCase(txtSecuCode), "''", "S")
					lgStrSQL = lgStrSQL & " AND SEQ = " & ArrColVal(3)
					
					

				End Select
				
			End If
		
		Next
		
		'Response.Write lgstrsql
	
	End If
'		Response.Write "SMMMMMMMMMMMMMMMMMMMMMM" & flg
    If lgStrSQL <> "" Then
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("SM" & flg,lgObjConn,lgObjRs,Err)
	END If



End Sub
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext
                  Case " "
                  Case "P"
                  Case "N"
           End Select
        Case "M"

           Select Case Mid(pDataType,2,1)
                  Case "R"
                    Select Case lgPrevNext
                        Case ""
                            lgStrSQL = "Select              a.security_cd, "
                            lgStrSQL = lgStrSQL & "         a.security_nm, "
                            lgStrSQL = lgStrSQL & "         a.temp_gl_dt as temp_gl_dt, "
                            lgStrSQL = lgStrSQL & "         a.security_type, "
                            lgStrSQL = lgStrSQL & "         b.minor_nm as security_type_nm, "
                            lgStrSQL = lgStrSQL & "         a.doc_cur, "
                            lgStrSQL = lgStrSQL & "         i.currency_desc, "
                            lgStrSQL = lgStrSQL & "         a.xch_rate, "
                            lgStrSQL = lgStrSQL & "         a.buy_amt, "
                            lgStrSQL = lgStrSQL & "         a.dept_cd1, "
                            lgStrSQL = lgStrSQL & "         e.dept_nm as dept_nm1, "
                            lgStrSQL = lgStrSQL & "         a.biz_area_cd, "
                            lgStrSQL = lgStrSQL & "         a.org_change_id, "
                            lgStrSQL = lgStrSQL & "         a.loc_buy_amt, "
                            lgStrSQL = lgStrSQL & "         a.dept_cd2, "
                            lgStrSQL = lgStrSQL & "         f.dept_nm as dept_nm2, "
                            lgStrSQL = lgStrSQL & "         a.internal_cd1, "
                            lgStrSQL = lgStrSQL & "         a.internal_cd2, "
                            lgStrSQL = lgStrSQL & "         a.price_amt, "
                            lgStrSQL = lgStrSQL & "         a.cust_cd1, "
                            lgStrSQL = lgStrSQL & "         g.bp_nm as bp_nm1, "
                            lgStrSQL = lgStrSQL & "         a.loc_price_amt, "
                            lgStrSQL = lgStrSQL & "         a.cust_cd2, "
                            lgStrSQL = lgStrSQL & "         h.bp_nm as bp_nm2, "
                            lgStrSQL = lgStrSQL & "         a.cnt, "
                            lgStrSQL = lgStrSQL & "         isnull(a.end_yn," & FilterVar("N", "''", "S") & " ) as end_yn, "
                            lgStrSQL = lgStrSQL & "         a.publ_dt, "
                            lgStrSQL = lgStrSQL & "         isnull(a.calcu_yn," & FilterVar("N", "''", "S") & " ) as calcu_yn, "
                            lgStrSQL = lgStrSQL & "         a.int_rate, "
                            lgStrSQL = lgStrSQL & "         isnull(a.cd_mtd," & FilterVar("1", "''", "S") & " ) as cd_mtd, "
                            lgStrSQL = lgStrSQL & "         a.expir_dt, "
                            lgStrSQL = lgStrSQL & "         a.ref_no, "
                            lgStrSQL = lgStrSQL & "         a.in_dt, "
                            lgStrSQL = lgStrSQL & "         a.temp_gl_no, "
                            lgStrSQL = lgStrSQL & "         a.out_dt, "
                            lgStrSQL = lgStrSQL & "         a.gl_no, "
                            lgStrSQL = lgStrSQL & "         a.acct_cd1, "
                            lgStrSQL = lgStrSQL & "         k.acct_nm as acct_nm1, "
                            lgStrSQL = lgStrSQL & "         a.acct_cd2, "
                            lgStrSQL = lgStrSQL & "         l.acct_nm as acct_nm2, "
                            lgStrSQL = lgStrSQL & "         a.org_change_id "
                            lgStrSQL = lgStrSQL & "   from  a_security a, "
                            lgStrSQL = lgStrSQL & "         b_minor b, "
                            lgStrSQL = lgStrSQL & "         b_acct_dept e, "
                            lgStrSQL = lgStrSQL & "         b_acct_dept f, "
                            lgStrSQL = lgStrSQL & "         b_biz_partner g, "
                            lgStrSQL = lgStrSQL & "         b_biz_partner h, "
                            lgStrSQL = lgStrSQL & "         b_currency i, "
                            lgStrSQL = lgStrSQL & "         b_cost_center j ,"
                            lgStrSQL = lgStrSQL & "         a_acct k, "
                            lgStrSQL = lgStrSQL & "         a_acct l "
                            lgStrSQL = lgStrSQL & "  where  a.security_type = b.minor_cd "
                            lgStrSQL = lgStrSQL & "    and  b.major_cd = " & FilterVar("A1031", "''", "S") & "  "
                            lgStrSQL = lgStrSQL & "    and  a.dept_cd1 = e.dept_cd "
                            lgStrSQL = lgStrSQL & "    and  e.org_change_id= (select cur_org_change_id from b_company) "
                            lgStrSQL = lgStrSQL & "    and  a.dept_cd2 = f.dept_cd "
                            lgStrSQL = lgStrSQL & "    and  f.org_change_id = (select cur_org_change_id from b_company) "
                            lgStrSQL = lgStrSQL & "    and  a.cust_cd1 = g.bp_cd "
                            lgStrSQL = lgStrSQL & "    and  a.cust_cd2 *= h.bp_cd "
                            lgStrSQL = lgStrSQL & "    and  a.doc_cur *= i.currency "
                            lgStrSQL = lgStrSQL & "    and  e.cost_cd = j.cost_cd "
                            lgStrSQL = lgStrSQL & "    and  a.acct_cd1 = k.acct_cd "
                            lgStrSQL = lgStrSQL & "    and  a.acct_cd2 = l.acct_cd "
                            lgStrSQL = lgStrSQL & "    and  a.security_cd = " & FilterVar(txtSecuCode, "''", "S")
                            lgStrSQL = lgStrSQL & "order by a.security_cd "
                            
                        Case "P"
                            lgStrSQL = "Select    top 1     a.security_cd, "
                            lgStrSQL = lgStrSQL & "         a.security_nm, "
                            lgStrSQL = lgStrSQL & "         a.temp_gl_dt as temp_gl_dt, "
                            lgStrSQL = lgStrSQL & "         a.security_type, "
                            lgStrSQL = lgStrSQL & "         b.minor_nm as security_type_nm, "
                            lgStrSQL = lgStrSQL & "         a.doc_cur, "
                            lgStrSQL = lgStrSQL & "         i.currency_desc, "
                            lgStrSQL = lgStrSQL & "         a.xch_rate, "
                            lgStrSQL = lgStrSQL & "         a.buy_amt, "
                            lgStrSQL = lgStrSQL & "         a.dept_cd1, "
                            lgStrSQL = lgStrSQL & "         e.dept_nm as dept_nm1, "
                            lgStrSQL = lgStrSQL & "         a.biz_area_cd, "
                            lgStrSQL = lgStrSQL & "         a.org_change_id, "
                            lgStrSQL = lgStrSQL & "         a.loc_buy_amt, "
                            lgStrSQL = lgStrSQL & "         a.dept_cd2, "
                            lgStrSQL = lgStrSQL & "         f.dept_nm as dept_nm2, "
                            lgStrSQL = lgStrSQL & "         a.internal_cd1, "
                            lgStrSQL = lgStrSQL & "         a.internal_cd2, "
                            lgStrSQL = lgStrSQL & "         a.price_amt, "
                            lgStrSQL = lgStrSQL & "         a.cust_cd1, "
                            lgStrSQL = lgStrSQL & "         g.bp_nm as bp_nm1, "
                            lgStrSQL = lgStrSQL & "         a.loc_price_amt, "
                            lgStrSQL = lgStrSQL & "         a.cust_cd2, "
                            lgStrSQL = lgStrSQL & "         h.bp_nm as bp_nm2, "
                            lgStrSQL = lgStrSQL & "         a.cnt, "
                            lgStrSQL = lgStrSQL & "         isnull(a.end_yn," & FilterVar("N", "''", "S") & " ) as end_yn, "
                            lgStrSQL = lgStrSQL & "         a.publ_dt, "
                            lgStrSQL = lgStrSQL & "         isnull(a.calcu_yn," & FilterVar("N", "''", "S") & " ) as calcu_yn, "
                            lgStrSQL = lgStrSQL & "         a.int_rate, "
                            lgStrSQL = lgStrSQL & "         isnull(a.cd_mtd," & FilterVar("1", "''", "S") & " ) as cd_mtd, "
                            lgStrSQL = lgStrSQL & "         a.expir_dt, "
                            lgStrSQL = lgStrSQL & "         a.ref_no, "
                            lgStrSQL = lgStrSQL & "         a.in_dt, "
                            lgStrSQL = lgStrSQL & "         a.temp_gl_no, "
                            lgStrSQL = lgStrSQL & "         a.out_dt, "
                            lgStrSQL = lgStrSQL & "         a.gl_no, "
                            lgStrSQL = lgStrSQL & "         a.acct_cd1, "
                            lgStrSQL = lgStrSQL & "         k.acct_nm as acct_nm1, "
                            lgStrSQL = lgStrSQL & "         a.acct_cd2, "
                            lgStrSQL = lgStrSQL & "         l.acct_nm as acct_nm2, "
                            lgStrSQL = lgStrSQL & "         a.org_change_id "
                            lgStrSQL = lgStrSQL & "   from  a_security a, "
                            lgStrSQL = lgStrSQL & "         b_minor b, "
                            lgStrSQL = lgStrSQL & "         b_acct_dept e, "
                            lgStrSQL = lgStrSQL & "         b_acct_dept f, "
                            lgStrSQL = lgStrSQL & "         b_biz_partner g, "
                            lgStrSQL = lgStrSQL & "         b_biz_partner h, "
                            lgStrSQL = lgStrSQL & "         b_currency i, "
                            lgStrSQL = lgStrSQL & "         b_cost_center j, "
                            lgStrSQL = lgStrSQL & "         a_acct k, "
                            lgStrSQL = lgStrSQL & "         a_acct l "
                            lgStrSQL = lgStrSQL & "  where  a.security_type = b.minor_cd "
                            lgStrSQL = lgStrSQL & "    and  b.major_cd = " & FilterVar("A1031", "''", "S") & "  "
                            lgStrSQL = lgStrSQL & "    and  a.dept_cd1 = e.dept_cd "
                            lgStrSQL = lgStrSQL & "    and  e.org_change_id = (select cur_org_change_id from b_company) "
                            lgStrSQL = lgStrSQL & "    and  a.dept_cd2 = f.dept_cd "
                            lgStrSQL = lgStrSQL & "    and  f.org_change_id = (select cur_org_change_id from b_company) "
                            lgStrSQL = lgStrSQL & "    and  a.cust_cd1 = g.bp_cd "
                            lgStrSQL = lgStrSQL & "    and  a.cust_cd2 *= h.bp_cd "
                            lgStrSQL = lgStrSQL & "    and  a.doc_cur *= i.currency "
                            lgStrSQL = lgStrSQL & "    and  e.cost_cd = j.cost_cd "
                            lgStrSQL = lgStrSQL & "    and  a.acct_cd1 = k.acct_cd "
                            lgStrSQL = lgStrSQL & "    and  a.acct_cd2 = l.acct_cd "
                            lgStrSQL = lgStrSQL & "    and  a.security_cd < " & FilterVar(txtSecuCode, "''", "S")
                            lgStrSQL = lgStrSQL & "order by a.security_cd desc"
                        Case "N"
                            lgStrSQL = "Select  top 1       a.security_cd, "
                            lgStrSQL = lgStrSQL & "         a.security_nm, "
                            lgStrSQL = lgStrSQL & "         a.temp_gl_dt as temp_gl_dt, "
                            lgStrSQL = lgStrSQL & "         a.security_type, "
                            lgStrSQL = lgStrSQL & "         b.minor_nm as security_type_nm, "
                            lgStrSQL = lgStrSQL & "         a.doc_cur, "
                            lgStrSQL = lgStrSQL & "         i.currency_desc, "
                            lgStrSQL = lgStrSQL & "         a.xch_rate, "
                             lgStrSQL = lgStrSQL & "         a.buy_amt, "
                            lgStrSQL = lgStrSQL & "         a.dept_cd1, "
                            lgStrSQL = lgStrSQL & "         e.dept_nm as dept_nm1, "
                            lgStrSQL = lgStrSQL & "         a.biz_area_cd, "
                            lgStrSQL = lgStrSQL & "         a.org_change_id, "
                            lgStrSQL = lgStrSQL & "         a.loc_buy_amt, "
                            lgStrSQL = lgStrSQL & "         a.dept_cd2, "
                            lgStrSQL = lgStrSQL & "         f.dept_nm as dept_nm2, "
                            lgStrSQL = lgStrSQL & "         a.internal_cd1, "
                            lgStrSQL = lgStrSQL & "         a.internal_cd2, "
                            lgStrSQL = lgStrSQL & "         a.price_amt, "
                            lgStrSQL = lgStrSQL & "         a.cust_cd1, "
                            lgStrSQL = lgStrSQL & "         g.bp_nm as bp_nm1, "
                            lgStrSQL = lgStrSQL & "         a.loc_price_amt, "
                            lgStrSQL = lgStrSQL & "         a.cust_cd2, "
                            lgStrSQL = lgStrSQL & "         h.bp_nm as bp_nm2, "
                            lgStrSQL = lgStrSQL & "         a.cnt, "
                            lgStrSQL = lgStrSQL & "         isnull(a.end_yn," & FilterVar("N", "''", "S") & " ) as end_yn, "
                            lgStrSQL = lgStrSQL & "         a.publ_dt, "
                            lgStrSQL = lgStrSQL & "         isnull(a.calcu_yn," & FilterVar("N", "''", "S") & " ) as calcu_yn, "
                            lgStrSQL = lgStrSQL & "         a.int_rate, "
                            lgStrSQL = lgStrSQL & "         isnull(a.cd_mtd," & FilterVar("1", "''", "S") & " ) as cd_mtd, "
                            lgStrSQL = lgStrSQL & "         a.expir_dt, "
                            lgStrSQL = lgStrSQL & "         a.ref_no, "
                            lgStrSQL = lgStrSQL & "         a.in_dt, "
                            lgStrSQL = lgStrSQL & "         a.temp_gl_no, "
                            lgStrSQL = lgStrSQL & "         a.out_dt, "
                            lgStrSQL = lgStrSQL & "         a.gl_no, "
                            lgStrSQL = lgStrSQL & "         a.acct_cd1, "
                            lgStrSQL = lgStrSQL & "         k.acct_nm as acct_nm1, "
                            lgStrSQL = lgStrSQL & "         a.acct_cd2, "
                            lgStrSQL = lgStrSQL & "         l.acct_nm as acct_nm2, "
                            lgStrSQL = lgStrSQL & "         a.org_change_id "
                            lgStrSQL = lgStrSQL & "   from  a_security a, "
                            lgStrSQL = lgStrSQL & "         b_minor b, "
                            lgStrSQL = lgStrSQL & "         b_acct_dept e, "
                            lgStrSQL = lgStrSQL & "         b_acct_dept f, "
                            lgStrSQL = lgStrSQL & "         b_biz_partner g, "
                            lgStrSQL = lgStrSQL & "         b_biz_partner h, "
                            lgStrSQL = lgStrSQL & "         b_currency i, "
                            lgStrSQL = lgStrSQL & "         b_cost_center j "
                            lgStrSQL = lgStrSQL & "         a_acct k, "
                            lgStrSQL = lgStrSQL & "         a_acct l "
                            lgStrSQL = lgStrSQL & "  where  a.security_type = b.minor_cd "
                            lgStrSQL = lgStrSQL & "    and  b.major_cd = " & FilterVar("A1031", "''", "S") & "  "
                            lgStrSQL = lgStrSQL & "    and  a.dept_cd1 = e.dept_cd "
                            lgStrSQL = lgStrSQL & "    and  e.org_change_id = (select cur_org_change_id from b_company) "
                            lgStrSQL = lgStrSQL & "    and  a.dept_cd2 = f.dept_cd "
                            lgStrSQL = lgStrSQL & "    and  f.org_change_id = (select cur_org_change_id from b_company) "
                            lgStrSQL = lgStrSQL & "    and  a.cust_cd1 = g.bp_cd "
                            lgStrSQL = lgStrSQL & "    and  a.cust_cd2 *= h.bp_cd "
                            lgStrSQL = lgStrSQL & "    and  a.doc_cur *= i.currency "
                            lgStrSQL = lgStrSQL & "    and  e.cost_cd = j.cost_cd "
                            lgStrSQL = lgStrSQL & "    and  a.acct_cd1 = k.acct_cd "
                            lgStrSQL = lgStrSQL & "    and  a.acct_cd2 = l.acct_cd "
                            lgStrSQL = lgStrSQL & "    and  a.security_cd > " & FilterVar(txtSecuCode, "''", "S")
                            lgStrSQL = lgStrSQL & "order by a.security_cd "
                    
                    End Select
               Case "U"
                       lgStrSQL = "UPDATE B_MAJOR  .......... "
			   Case "M"		'Multi
						
						lgStrSQL = " Select A.SEQ,A.PAY_TYPE, C.MINOR_NM,isnull(A.PAY_AMT,0) as PAY_AMT,isnull(A.PAY_LOC_AMT,0) as PAY_LOC_AMT,isnull(A.NOTE_NO,'') as NOTE_NO,isnull(A.BANK_NO,'') as BANK_NO,isnull(B.BANK_NM,'') AS BANK_NM,isnull(A.BANK_ACCT_NO,'') AS BANK_ACCT_NO "
						lgStrSQL = lgStrSQL & " From A_SECURITY_ITEM A,B_BANK B, B_MINOR C" 
						lgStrSQL = lgStrSQL & " Where A.SECURITY_CD = " & pCode  & " AND A.BANK_NO *= B.BANK_CD AND A.PAY_TYPE = C.MINOR_CD AND C.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " "  

						'lgStrSQL = " Select SEQ,PAY_TYPE,isnull(A.PAY_AMT,0) as PAY_AMT,isnull(A.PAY_LOC_AMT,0) as PAY_LOC_AMT,isnull(A.NOTE_NO,'') as NOTE_NO,isnull(A.BANK_NO,'') as BANK_NO,isnull(B.BANK_NM,'') AS BANK_NM,isnull(A.BANK_ACCT_NO,'') AS BANK_ACCT_NO "
						'lgStrSQL = lgStrSQL & " From A_SECURITY_ITEM A,B_BANK B " 
						'lgStrSQL = lgStrSQL & " Where A.SECURITY_CD = " & pCode  & " AND A.BANK_NO *= B.BANK_CD"  
						
						
           End Select
		 Case "Q"
			lgStrSQL = "SELECT TOP 1 SECURITY_CD "
            lgStrSQL = lgStrSQL & " From  A_SECURITY "
            lgStrSQL = lgStrSQL & " WHERE SECURITY_CD = " & pCode	       
    End Select
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
    ObjectContext.SetAbort                                                        '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
	
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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
               With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>", "F"                          '☜ : Display data
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
				Call .ReFormatSpreadCellByCellByCurrency2(.Frm1.vspdData,"<%=lgLngMaxRow +1 %>" , "<%= lgLngMaxRow + iLoopCount%>" ,.frm1.txtTradeCur.value ,.C_AMT ,   "A" ,"I","X","X")

				.DBQueryOk
 	         End with
         End If
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else
          End If
    End Select


</Script>
