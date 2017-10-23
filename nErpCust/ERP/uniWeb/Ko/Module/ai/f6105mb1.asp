<%
'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : Prepayment
'*  3. Program ID           : f6105mb1
'*  4. Program Name         : 선급금 기초등록 
'*  5. Program Desc         : 선급금 기초등록 
'*  6. Modified date(First) : 2000/09/27
'*  7. Modified date(Last)  : 2002/11/18
'*  8. Modifier (First)     : 조익성 
'*  9. Modifier (Last)      : Jeong Yong Kyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Expires = -1														'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True														'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd

On Error Resume Next														'☜: 
Err.Clear																	'☜: Protect system from crashing

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim iPAFG605																'입력/수정용 ComProxy Dll 사용 변수 
Dim istrCode																'Lookup 용 코드 저장 변수 
Dim istrMode																'현재 MyBiz.asp 의 진행상태를 나타냄 
Dim iCommandSent

Dim iImp_prpaym
Const C_IMP_PRPAYM_NO = 0
Const C_IMP_PRPAYM_DT = 1
Const C_IMP_PRPAYM_TYPE = 2
Const C_IMP_PAYM_TYPE = 3
Const C_IMP_REF_NO = 4
Const C_IMP_DOC_CUR = 5
Const C_IMP_XCH_RATE = 6
Const C_IMP_NOTE_NO = 7
Const C_IMP_PRPAYM_AMT = 8
Const C_IMP_PRPAYM_LOC_AMT = 9
Const C_IMP_BAL_AMT = 10
Const C_IMP_BAL_LOC_AMT = 11
Const C_IMP_GL_NO = 12
Const C_IMP_TEMP_GL_NO = 13
Const C_IMP_PRPAYM_STS = 14
Const C_IMP_CONF_FG = 15
Const C_IMP_PRPAYM_FG = 16
Const C_IMP_VAT_TYPE = 17
Const C_IMP_VAT_AMT = 18
Const C_IMP_VAT_LOC_AMT = 19
Const C_IMP_PRPAYM_DESC = 20
Const C_IMP_ISSUED_DT = 21
Const C_IMP_Gl_DT = 22
    
Dim iImp_dept
Const C_IMP_ORG_CHANGE_ID = 0
Const C_IMP_DEPT_CD = 1

Dim iImp_b_bank_acct
Dim iImp_b_bank
Dim iImp_biz_partner
Dim iImp_b_currency
Dim iImp_a_acct_trans_type

istrMode = Request("txtMode")													'현재 상태를 받음 

if Request("SelectChar")="" or isnull(Request("SelectChar")) then
	iCommandSent = "QUERY"
Else		
	iCommandSent = Request("SelectChar")
End if		

Select Case istrMode
    Case CStr(UID_M0001)                                                         '☜: Query
         Call SubBizQuery()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSave()
    Case CStr(UID_M0003)                                                         '☜: Delete
         Call SubBizDelete()
End Select

Sub SubBizQuery()

	On Error Resume Next
	Err.Clear 
    
	Dim lgCurrency
    Dim iarrPrpaym 
    Const C_PRPAYM_NO = 0
    Const C_PRPAYM_FG = 1
    Const C_ORG_CHANGE_ID = 2
    
    Dim iRarrPrpaym
    Const C_EXP_PRPAYM_NO = 0
    Const C_EXP_PRPAYM_TYPE = 1
    Const C_EXP_PAYM_TYPE_NM = 2
    Const C_EXP_DEPT_CD = 3
    Const C_EXP_DEPT_NM = 4
    Const C_EXP_BP_CD = 5
    Const C_EXP_BP_NM = 6
    Const C_EXP_PRPAYM_DT = 7
    Const C_EXP_JNL_CD = 8
    Const C_EXP_JNL_NM = 9
    Const C_EXP_NOTE_NO = 10
    Const C_EXP_BANK_CD = 11
    Const C_EXP_BANK_NM = 12
    Const C_EXP_BANK_ACCT_NO = 13
    Const C_EXP_DOC_CUR = 14
    Const C_EXP_XCH_RATE = 15
    Const C_EXP_PRPAYM_AMT = 16
    Const C_EXP_PRPAYM_LOC_AMT = 17
    Const C_EXP_CLS_AMT = 18
    Const C_EXP_CLS_LOC_AMT = 19
    Const C_EXP_STTL_AMT = 20
    Const C_EXP_STTL_LOC_AMT = 21
    Const C_EXP_BAL_AMT = 22
    Const C_EXP_BAL_LOC_AMT = 23
    Const C_EXP_VAT_TYPE = 24
    Const C_EXP_VAT_TYPE_NM = 25
    Const C_EXP_VAT_AMT = 26
    Const C_EXP_VAT_LOC_AMT = 27
    Const C_EXP_GL_NO = 28
    Const C_EXP_TEMP_GL_NO = 29
    Const C_EXP_REF_NO = 30
    Const C_EXP_PRPAYM_DESC = 31
    Const C_EXP_IO_FG = 32
    Const C_EXP_IO_FG_NM = 33    
    Const C_EXP_TAX_BIZ_AREA_CD = 34
    Const C_EXP_TAX_BIZ_AREA_NM = 35
    Const C_EXP_ISSUED_DT = 36    
    Const C_EXP_ACCT_CD = 37
    Const C_EXP_ACCT_NM = 38        
    Const C_EXP_GL_Dt = 39 
    
    Set iPAFG605 = Server.CreateObject("PAFG605.cFLkUpPpSvr")

    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If CheckSYSTEMError(Err, True) = True Then					
		Response.End 
	End If   
    
    '------------------------------------------
    'Data manipulate  area(import view match)
    '------------------------------------------
	Redim iArrPrpaym(2)
	
	iArrPrpaym(C_PRPAYM_NO) = Trim(Request("txtPrpaymNo"))
	iArrPrpaym(C_PRPAYM_FG) = "PT"
	iArrPrpaym(C_ORG_CHANGE_ID) = GetGlobalInf("gChangeOrgId")
	
    '------------------------------------------
    'Com action area
    '------------------------------------------
	
	Call iPAFG605.F_LOOKUP_PRPAYM_SVR(gStrGloBalCollection,iCommandSent,iarrPrpaym,iRarrPrpaym)

	'------------------------------------------
	'Com action result check area(OS,internal)
	'------------------------------------------
	if err.number <> "0" then
		If iCommandSent = "QUERY" Then
		   Call DisplayMsgBox("141500", vbInformation, "", "", I_MKSCRIPT)            '☜ : No data is found. 
		ElseIf iCommandSent = "PREV" Then
		   Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)            '☜ : This is the starting data. 
		   iCommandSent = "QUERY"
		   Call SubBizQuery()
		ElseIf iCommandSent = "NEXT" Then
		   Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)            '☜ : This is the ending data.
		   iCommandSent = "QUERY"
		   Call SubBizQuery()
		End If
	Else    
		lgCurrency = ConvSPChars(iRarrPrpaym(C_EXP_DOC_CUR))

		Response.Write "<Script Language=vbscript>"                                                              & vbCr
		Response.Write "With parent.frm1"															   	         & vbCr
		Response.Write ".txtPrpaymNo.value      = """ & ConvSPChars(iRarrPrpaym(C_EXP_PRPAYM_NO))         & """" & vbCr
		Response.Write ".txtPrpaymType.value    = """ & ConvSPChars(iRarrPrpaym(C_EXP_PRPAYM_TYPE))       & """" & vbCr
		Response.Write ".txtPrpaymTypeNm.value  = """ & ConvSPChars(iRarrPrpaym(C_EXP_PAYM_TYPE_NM))    & """" & vbCr
			
		Response.Write ".txtDeptCd.value        = """ & ConvSPChars(iRarrPrpaym(C_EXP_DEPT_CD))           & """" & vbCr
		Response.Write ".txtDeptNm.value        = """ & ConvSPChars(iRarrPrpaym(C_EXP_DEPT_NM))           & """" & vbCr
		Response.Write ".txtPrpaymDt.Text       = """ & UNIDateClientFormat(iRarrPrpaym(C_EXP_PRPAYM_DT)) & """" & vbCr
		Response.Write ".txtBpCd.value          = """ & ConvSPChars(iRarrPrpaym(C_EXP_BP_CD))             & """" & vbCr
		Response.Write ".txtBpNm.value          = """ & ConvSPChars(iRarrPrpaym(C_EXP_BP_NM))             & """" & vbCr
		Response.Write ".txtDocCur.value        = """ & ConvSPChars(iRarrPrpaym(C_EXP_DOC_CUR))           & """" & vbCr
		Response.Write ".txtPrpaymAmt.Text      = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_PRPAYM_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                & """" & vbCr
		Response.Write ".txtClsAmt.Text         = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_CLS_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                   & """" & vbCr
		Response.Write ".txtSttlAmt.Text        = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_STTL_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                  & """" & vbCr
		Response.Write ".txtBalAmt.Text         = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_BAL_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                   & """" & vbCr

		If gIsShowLocal <> "N" Then				
			Response.Write ".txtXchRate.Text        = """ & UNINumClientFormat(iRarrPrpaym(C_EXP_XCH_RATE), ggExchRate.DecPoint, 0)                                             & """" & vbCr
			Response.Write ".txtPrpaymLocAmt.Text   = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_PRPAYM_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr
			Response.Write ".txtClsLocAmt.Text	    = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_CLS_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr
			Response.Write ".txtSttlLocAmt.Text     = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_STTL_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")    & """" & vbCr
			Response.Write ".txtBalLocAmt.Text      = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_BAL_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr		
		Else
			Response.Write ".txtXchRate.Value       = """ & UNINumClientFormat(iRarrPrpaym(C_EXP_XCH_RATE), ggExchRate.DecPoint, 0)                                             & """" & vbCr
			Response.Write ".txtPrpaymLocAmt.Value  = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_PRPAYM_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr
			Response.Write ".txtClsLocAmt.Value	    = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_CLS_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr
			Response.Write ".txtSttlLocAmt.Value    = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_STTL_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")    & """" & vbCr
			Response.Write ".txtBalLocAmt.Value     = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_BAL_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr		
		End if		
		C_EXP_GL_Dt
		Response.Write ".txtTempGlNo.value      = """ & ConvSPChars(iRarrPrpaym(C_EXP_TEMP_GL_NO))        & """" & vbCr
		Response.Write ".txtGlNo.value          = """ & ConvSPChars(iRarrPrpaym(C_EXP_GL_NO))             & """" & vbCr
		Response.Write ".txtPrpaymDesc.value    = """ & ConvSPChars(iRarrPrpaym(C_EXP_PRPAYM_DESC))       & """" & vbCr
		Response.Write ".txtGlDt.Text       = """ & UNIDateClientFormat(iRarrPrpaym(C_EXP_GL_Dt)) & """" & vbCr
		
		Response.Write "parent.DbQueryOk          "                                                              & vbCr 															'☜: 조회가 성공 
		Response.Write "End With                  "                                                              & vbCr
		Response.Write "</Script>                 "                                                              & vbCr
	End if

    Set iPAFG605 = Nothing															'☜: Unload Complus

End Sub

Sub SubBizSave()
									
	On Error Resume Next
	Err.Clear 	
	
	Dim iE_prpaym_no 
	
    If Request("txtFlgMode") = "" Then				                                '저장을 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)	            'txtFlgMode 조건값이 비어있습니다!
		Response.End 
	End If
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

    Set iPAFG605 = Server.CreateObject("PAFG605.cFMngPpSvr")
    
    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If CheckSYSTEMError(Err, True) = True Then					
		Response.End 
	End If   
	
    '------------------------------------------
    'Data manipulate area
    '------------------------------------------
	Redim iImp_prpaym(C_IMP_Gl_DT)
	Redim iImp_dept(1)

	iImp_prpaym(C_IMP_PRPAYM_NO)	  =  UCase(Trim(Request("txtPrpaymNo")))
	iImp_prpaym(C_IMP_PRPAYM_DT)	  =  UNIConvDate(Request("txtPrpaymDt"))
	iImp_prpaym(C_IMP_PRPAYM_TYPE)	  =  Trim(UCase(Request("txtPrpaymType")))
	iImp_prpaym(C_IMP_DOC_CUR)        =  Trim(UCase(Request("txtDocCur")))
	iImp_prpaym(C_IMP_XCH_RATE)		  =  UNIConvNum(Request("txtXchRate"),0)
	iImp_prpaym(C_IMP_PRPAYM_AMT)	  =  UNIConvNum(Request("txtPrpaymAmt"),0)
	iImp_prpaym(C_IMP_PRPAYM_LOC_AMT) =  UNIConvNum(Request("txtPrpaymLocAmt"),0)
	iImp_prpaym(C_IMP_BAL_AMT)		  =  UNIConvNum(Request("txtBalAmt"),0)
	iImp_prpaym(C_IMP_BAL_LOC_AMT)	  =  UNIConvNum(Request("txtBalLocAmt"),0)
	iImp_prpaym(C_IMP_GL_NO)          =  Trim(Request("txtGlNo"))
	iImp_prpaym(C_IMP_TEMP_GL_NO)     =  Trim(Request("txtTempGlNo"))
	iImp_prpaym(C_IMP_PRPAYM_STS)     =  ""		
	iImp_prpaym(C_IMP_CONF_FG)        =  ""
	iImp_prpaym(C_IMP_PRPAYM_FG)	  =  "PT"
	iImp_prpaym(C_IMP_PRPAYM_DESC)	  =  Trim(Request("txtPrpaymDesc"))
	iImp_prpaym(C_IMP_Gl_DT)		  =  UNIConvDate(Request("txtGlDt"))
	
	
	iImp_dept(C_IMP_ORG_CHANGE_ID)    =  GetGlobalInf("gChangeOrgId")
	iImp_dept(C_IMP_DEPT_CD)          =  UCase(Trim(Request("txtDeptCd")))
	
	iImp_biz_partner	              =  UCase(Trim(Request("txtBpCd")))
	iImp_b_currency                   =  gCurrency	
	iImp_a_acct_trans_type            = "FP003"
    
	'------------------------------------------
	'Com Action Area
	'------------------------------------------

    If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
    End If
    
    iE_prpaym = iPAFG605.F_MANAGE_PRPAYM_SVR(gStrGloBalCollection,iCommandSent,iImp_prpaym,iImp_dept,, _
		                           ,iImp_biz_partner,iImp_b_currency,iImp_a_acct_trans_type)
    
  	if CheckSYSTEMError(Err, True) = True Then					
		Set iPAFG605 = Nothing
		Response.End 
	End If  
		
	'------------------------------------------
	'Result data display area
	'------------------------------------------
    Response.Write "<Script Language=vbscript>    "                                & vbCr
	Response.Write "With parent                   "                                & vbCr
    Response.write " If .frm1.txtPrpaymNo.Value = """"   Then "					   & vbCr
    Response.Write " .frm1.txtPrpaymNo.Value = """ & ConvSPChars(iE_prpaym)	& """" & vbCr
	Response.write " End If "													   & vbCr                       
	Response.Write ".DbSaveOk                     "                                & vbCr
	Response.Write "End With                      "                                & vbCr
	Response.Write "</Script>                     "                                & vbCr

    Set iPAFG605 = Nothing															'☜: Unload Complus

End Sub
	    
Sub SubBizDelete()

	On Error Resume Next
	Err.Clear 

    If Request("txtPrpaymNo") = "" Then												'삭제를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("900002", vbInformation, "", "", I_MKSCRIPT)				'조회를 먼저 하세요.
		Response.End 
	End If
    
    Set iPAFG605 = Server.CreateObject("PAFG605.cFMngPpSvr")

    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If CheckSYSTEMError(Err, True) = True Then					
		Response.End 
	End If   
    
    '------------------------------------------
    'Data manipulate area
    '------------------------------------------
    Redim iImp_prpaym(C_IMP_Gl_DT)
	Redim iImp_dept(1)
    
    iCommandSent = "DELETE"
		
	iImp_prpaym(C_IMP_PRPAYM_NO)	  = UCase(Trim(Request("txtPrpaymNo")))
	iImp_prpaym(C_IMP_PRPAYM_DT)	  = UNIConvDate(Request("txtPrpaymDt"))
	iImp_prpaym(C_IMP_PRPAYM_TYPE)	  = Trim(UCase(Request("txtPrpaymType")))
	iImp_prpaym(C_IMP_DOC_CUR)        = Trim(UCase(Request("txtDocCur")))
	iImp_prpaym(C_IMP_XCH_RATE)		  = UNIConvNum(Request("txtXchRate"),0)
	iImp_prpaym(C_IMP_PRPAYM_AMT)	  = UNIConvNum(Request("txtPrpaymAmt"),0)
	iImp_prpaym(C_IMP_PRPAYM_LOC_AMT) = UNIConvNum(Request("txtPrpaymLocAmt"),0)
	iImp_prpaym(C_IMP_BAL_AMT)		  = UNIConvNum(Request("txtBalAmt"),0)
	iImp_prpaym(C_IMP_BAL_LOC_AMT)	  = UNIConvNum(Request("txtBalLocAmt"),0)
	iImp_prpaym(C_IMP_GL_NO)          = Trim(Request("txtGlNo"))
	iImp_prpaym(C_IMP_TEMP_GL_NO)     = Trim(Request("txtTempGlNo"))
	iImp_prpaym(C_IMP_PRPAYM_STS)     = ""		
	iImp_prpaym(C_IMP_CONF_FG)        = ""
	iImp_prpaym(C_IMP_PRPAYM_FG)	  = "PT"
	iImp_prpaym(C_IMP_PRPAYM_DESC)	  = Trim(Request("txtPrpaymDesc"))
    iImp_prpaym(C_IMP_Gl_DT)			= UNIConvDate(Request("txtGlDt"))
    
    iImp_dept(C_IMP_ORG_CHANGE_ID)    = GetGlobalInf("gChangeOrgId")
	iImp_dept(C_IMP_DEPT_CD)          = UCase(Trim(Request("txtDeptCd")))
	
	iImp_biz_partner	              = UCase(Trim(Request("txtBpCd")))
	iImp_b_currency                   = gCurrency	
	iImp_a_acct_trans_type            = "FP003"
    
    '------------------------------------------
    'Com Action Area
    '------------------------------------------
	Call iPAFG605.F_MANAGE_PRPAYM_SVR(gStrGloBalCollection,iCommandSent,iImp_prpaym,iImp_dept,, _
		                           ,iImp_biz_partner,iImp_b_currency,iImp_a_acct_trans_type)
        
    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
  	if CheckSYSTEMError(Err, True) = True Then					
		Set iPAFG605 = Nothing
		Response.End 
	End If  

	'------------------------------------------
	'Result data display area
	'------------------------------------------

	Response.Write " <Script Language=vbscript>" & vbCr
	Response.Write " Call parent.DbDeleteOk() " & vbCr
	Response.Write " </Script> "			     & vbCr
					
    Set iPAFG605 = Nothing																'☜: Unload Complus

End Sub

%>
