<%
Option Explicit
'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : Prepayment
'*  3. Program ID           : f6101mb1
'*  4. Program Name         : 선급금 등록 
'*  5. Program Desc         : 선급금 등록 
'*  6. Modified date(First) : 2000/09/27
'*  7. Modified date(Last)  : 2002/11/15
'*  8. Modifier (First)     : 조익성 
'*  9. Modifier (Last)      : Jeong Yong Kyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
													'☜ : ASP가 캐쉬되지 않도록 한다.
													'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%																		'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd

On Error Resume Next													'☜: 
Err.Clear                                                               '☜: Protect system from crashing

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

Dim iPAFG605															'입력/수정용 ComProxy Dll 사용 변수 
Dim istrMode															'현재 MyBiz.asp 의 진행상태를 나타냄 
Dim iEG_Spread
Dim lgIntFlgMode
Dim iErrorPosition
Dim gIsShowLocal
Dim iImp_prpaym
Dim LngMaxRow
Dim LngRow
Dim strData
Dim lgStrPrevKey
Dim StrNextKey	
Dim iCommandSent

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

Const C_SPREAD_SEQ = 0
Const C_SPREAD_PAYM_TYPE =1
Const C_SPREAD_PAYM_TYPE_NM = 2
Const C_SPREAD_AMT = 3
Const C_SPREAD_LOC_AMT = 4
Const C_SPREAD_BANK_CD = 5
Const C_SPREAD_BANK_NM = 6 
Const C_SPREAD_BANK_ACCT_NO = 7
Const C_SPREAD_NOTE_NO = 8
Const C_SPREAD_ACCT_CD = 9
Const C_SPREAD_ACCT_NM = 10	
Const C_SPREAD_C_STTL_DESC = 11	

Dim iImp_dept
Const C_IMP_ORG_CHANGE_ID = 0
Const C_IMP_DEPT_CD = 1

Dim iImp_b_bank_acct
Dim iImp_b_bank
Dim iImp_biz_partner
Dim iImp_b_currency
Dim iImp_a_acct_trans_type
Dim iImp_a_acct
Dim iImp_b_tax_biz_area

istrMode = Request("txtMode")													 '현재 상태를 받음 

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
	Dim LookUp_states_char 

    Dim iarrPrpaym 
    Const C_PRPAYM_NO = 0
    Const C_PRPAYM_FG = 1
    Const C_ORG_CHANGE_ID = 2
    
    Dim iRarrPrpaym
    Const C_EXP_PRPAYM_NO = 0
    Const C_EXP_PRPAYM_TYPE = 1
    Const C_EXP_PRPAYM_TYPE_NM = 2
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

	' -- 권한관리추가 
	Const A745_I2_a_data_auth_data_BizAreaCd = 0
	Const A745_I2_a_data_auth_data_internal_cd = 1
	Const A745_I2_a_data_auth_data_sub_internal_cd = 2
	Const A745_I2_a_data_auth_data_auth_usr_id = 3

	Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A745_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I2_a_data_auth(A745_I2_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I2_a_data_auth(A745_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I2_a_data_auth(A745_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

    
    Set iPAFG605 = Server.CreateObject("PAFG605.cFLkUpPpSvr")

    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If CheckSYSTEMError(Err, True) = True Then					
		Exit Sub
	End If   

    '------------------------------------------
    'Data manipulate  area(import view match)
    '------------------------------------------
	Redim iArrPrpaym(2)
	
	iArrPrpaym(C_PRPAYM_NO) = Trim(Request("txtPrpaymNo"))
	iArrPrpaym(C_PRPAYM_FG) = "PP"
	iArrPrpaym(C_ORG_CHANGE_ID) = Trim(request("hOrgChangeId"))
	
	iCommandSent = Request("txtCommand")	
    '------------------------------------------
    'Com action area
    '------------------------------------------
	LookUp_states_char = iPAFG605.F_LOOKUP_PRPAYM_SVR(gStrGloBalCollection,iCommandSent,iarrPrpaym,iRarrPrpaym,iEG_Spread,I2_a_data_auth)

	'------------------------------------------
	'Com action result check area(OS,internal)
	'------------------------------------------
	If CheckSYSTEMError(Err, True) = True Then					
		Set iPAFG605 = Nothing
		Response.End 
	End If  

	LngMaxRow = Request("txtMaxRows")
	lgCurrency = ConvSPChars(iRarrPrpaym(C_EXP_DOC_CUR))

	Response.Write "<Script Language=vbscript>"                                                              & vbCr
	Response.Write "With parent.frm1"															   	         & vbCr
	Response.Write ".txtPrpaymNo.value      = """ & ConvSPChars(iRarrPrpaym(C_EXP_PRPAYM_NO))         & """" & vbCr
	Response.Write ".txtPrpaymType.value    = """ & ConvSPChars(iRarrPrpaym(C_EXP_PRPAYM_TYPE))       & """" & vbCr
	Response.Write ".txtPrpaymTypeNm.value  = """ & ConvSPChars(iRarrPrpaym(C_EXP_PRPAYM_TYPE_NM))    & """" & vbCr
			
	Response.Write ".txtPrpaymDt.text       = """ & UNIDateClientFormat(iRarrPrpaym(C_EXP_PRPAYM_DT)) & """" & vbCr
	Response.Write ".txtDeptCd.value        = """ & ConvSPChars(iRarrPrpaym(C_EXP_DEPT_CD))           & """" & vbCr
	Response.Write ".txtDeptNm.value        = """ & ConvSPChars(iRarrPrpaym(C_EXP_DEPT_NM))           & """" & vbCr
	Response.Write ".txtBpCd.value          = """ & ConvSPChars(iRarrPrpaym(C_EXP_BP_CD))             & """" & vbCr
	Response.Write ".txtBpNm.value          = """ & ConvSPChars(iRarrPrpaym(C_EXP_BP_NM))             & """" & vbCr
	Response.Write ".txtDocCur.value        = """ & ConvSPChars(iRarrPrpaym(C_EXP_DOC_CUR))           & """" & vbCr
	Response.Write ".txtRefNo.value	        = """ & ConvSPChars(iRarrPrpaym(C_EXP_REF_NO))            & """" & vbCr
	Response.Write ".txtPrpaymAmt.Text      = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_PRPAYM_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                & """" & vbCr
	Response.Write ".txtClsAmt.Text         = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_CLS_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                   & """" & vbCr
	Response.Write ".txtSttlAmt.Text        = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_STTL_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                  & """" & vbCr
	Response.Write ".txtBalAmt.Text         = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_BAL_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                   & """" & vbCr
	Response.Write ".txtVatAmt.Text         = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_VAT_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                   & """" & vbCr

	If gIsShowLocal <> "N" Then	
		Response.Write ".txtXchRate.Text        = """ & UNINumClientFormat(iRarrPrpaym(C_EXP_XCH_RATE),ggExchRate.DecPoint, 0)		                                        & """" & vbCr
		Response.Write ".txtPrpaymLocAmt.Text   = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_PRPAYM_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr
		Response.Write ".txtClsLocAmt.Text	    = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_CLS_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr
		Response.Write ".txtSttlLocAmt.Text     = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_STTL_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")    & """" & vbCr
		Response.Write ".txtBalLocAmt.Text      = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_BAL_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr
		Response.Write ".txtVatLocAmt.Text 	    = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_VAT_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr
	Else
		Response.Write ".txtXchRate.Value       = """ & UNINumClientFormat(iRarrPrpaym(C_EXP_XCH_RATE),ggExchRate.DecPoint, 0)		                                        & """" & vbCr
		Response.Write ".txtPrpaymLocAmt.Value  = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_PRPAYM_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr
		Response.Write ".txtClsLocAmt.Value	    = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_CLS_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr
		Response.Write ".txtSttlLocAmt.Value    = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_STTL_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")    & """" & vbCr
		Response.Write ".txtBalLocAmt.Value     = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_BAL_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr
		Response.Write ".txtVatLocAmt.Value     = """ & UNIConvNumDBToCompanyByCurrency(iRarrPrpaym(C_EXP_VAT_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr
	End if								
		
	Response.Write ".txtVatType.value           = """ & ConvSPChars(iRarrPrpaym(C_EXP_VAT_TYPE))          & """" & vbCr
	Response.Write ".txtVatTypeNm.value         = """ & ConvSPChars(iRarrPrpaym(C_EXP_VAT_TYPE_NM))       & """" & vbCr		
	Response.Write ".txtIssuedDt.text			= """ & UNIDateClientFormat(iRarrPrpaym(C_EXP_issued_dt)) & """" & vbCr		
	Response.Write ".txtBizAreaCD.value			= """ & ConvSPChars(iRarrPrpaym(C_EXP_tax_biz_area_cd))   & """" & vbCr 
	Response.Write ".txtBizAreaNM.value			= """ & ConvSPChars(iRarrPrpaym(C_EXP_tax_biz_area_nm))   & """" & VbCr			
	Response.Write ".txtTempGlNo.value          = """ & ConvSPChars(iRarrPrpaym(C_EXP_TEMP_GL_NO))        & """" & vbCr
	Response.Write ".txtGlNo.value              = """ & ConvSPChars(iRarrPrpaym(C_EXP_GL_NO))             & """" & vbCr
	Response.Write ".txtPrpaymDesc.value        = """ & ConvSPChars(iRarrPrpaym(C_EXP_PRPAYM_DESC))       & """" & vbCr
	Response.Write ".txtAfterLookUp.value		= """ & ConvSPChars(LookUp_states_char)                   & """" & vbCr		
'		Response.Write ConvSPChars(LookUp_states_char) 
	'LngMaxRow = parent.vspdData.MaxRows	

	For LngRow = 0 To UBOUND(iEG_Spread,1)
	    strData = strData & Chr(11) & iEG_Spread(LngRow,C_SPREAD_SEQ)
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_PAYM_TYPE))
		strData = strData & Chr(11) & " "    'popup button
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_PAYM_TYPE_NM))
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_ACCT_CD))
		strData = strData & Chr(11) & " "    'popup button
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_ACCT_NM))		
				
		strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iEG_Spread(LngRow,C_SPREAD_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iEG_Spread(LngRow,C_SPREAD_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
				
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_BANK_CD))
		strData = strData & Chr(11) & " "    'popup button
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_BANK_NM))
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_BANK_ACCT_NO))
		strData = strData & Chr(11) & " "    'popup button
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_NOTE_NO))
		strData = strData & Chr(11) & " "    'popup button
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_C_STTL_DESC))
		strData = strData & Chr(11) & LngMaxRow + LngRow + 1
	    strData = strData & Chr(11) & Chr(12)
	Next
	
    Response.Write "parent.ggoSpread.Source = .vspdData "					   & vbCr
 	Response.Write "parent.ggoSpread.SSShowData """ & strData			& """" & vbCr
    Response.Write "parent.lgStrPrevKey = """ & ConvSPChars(StrNextKey) & """" & vbCr 
    Response.Write "parent.DbQueryOk "                                         & vbCr
    Response.Write " End With  "                                               & vbCr
    Response.Write " </Script> "                                               & vbCr

    Set iPAFG605 = Nothing															'☜: Unload Complus
End Sub																		

Sub SubBizSave()

	On Error Resume Next
	Err.Clear 

	Dim iE_prpaym
	
    If Request("txtFlgMode") = "" Then				                                '저장을 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)	            'txtFlgMode 조건값이 비어있습니다!
		Exit Sub
	End If

	' -- 권한관리추가 
	Const A750_I10_a_data_auth_data_BizAreaCd = 0
	Const A750_I10_a_data_auth_data_internal_cd = 1
	Const A750_I10_a_data_auth_data_sub_internal_cd = 2
	Const A750_I10_a_data_auth_data_auth_usr_id = 3

	Dim I10_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I10_a_data_auth(3)
	I10_a_data_auth(A750_I10_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I10_a_data_auth(A750_I10_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I10_a_data_auth(A750_I10_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I10_a_data_auth(A750_I10_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
		
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

    Set iPAFG605 = Server.CreateObject("PAFG605.cFMngPpSvr")
    
    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If CheckSYSTEMError(Err, True) = True Then					
		Exit Sub
	End If   
	
    '------------------------------------------
    'Data manipulate area
    '------------------------------------------
	Redim iImp_prpaym(21)
	Redim iImp_dept(1)

    iImp_prpaym(C_IMP_PRPAYM_NO)	  =  UCase(Trim(Request("txtPrpaymNo")))
	iImp_prpaym(C_IMP_PRPAYM_DT)	  =  UNIConvDate(Request("txtPrpaymDt"))
	iImp_prpaym(C_IMP_PRPAYM_TYPE)	  =  Trim(UCase(Request("txtPrpaymType")))
	iImp_prpaym(C_IMP_PAYM_TYPE)	  =  UCase(Trim(Request("txtPaymType")))
	iImp_prpaym(C_IMP_REF_NO)         =  Trim(Request("txtRefNo"))
	iImp_prpaym(C_IMP_DOC_CUR)        =  UCase(Request("txtDocCur"))
	iImp_prpaym(C_IMP_XCH_RATE)		  =  UNIConvNum(Request("txtXchRate"),0)
'	iImp_prpaym(C_IMP_NOTE_NO)		  =  UCase(Trim(Request("txtNoteNo")))
	iImp_prpaym(C_IMP_PRPAYM_AMT)	  =  UNIConvNum(Request("txtPrpaymAmt"),0)
	iImp_prpaym(C_IMP_PRPAYM_LOC_AMT) =  UNIConvNum(Request("txtPrpaymLocAmt"),0)
	iImp_prpaym(C_IMP_BAL_AMT)		  =  UNIConvNum(Request("txtBalAmt"),0)
	iImp_prpaym(C_IMP_BAL_LOC_AMT)	  =  UNIConvNum(Request("txtBalLocAmt"),0)
	iImp_prpaym(C_IMP_GL_NO)          =  Trim(Request("txtGlNo"))
	iImp_prpaym(C_IMP_TEMP_GL_NO)     =  Trim(Request("txtTempGlNo"))
	iImp_prpaym(C_IMP_PRPAYM_STS)     =  ""		
	iImp_prpaym(C_IMP_CONF_FG)        =  ""
	iImp_prpaym(C_IMP_PRPAYM_FG)	  =  "PP"
	iImp_prpaym(C_IMP_VAT_TYPE)		  =  UCase(Trim(Request("txtVatType")))
	iImp_prpaym(C_IMP_VAT_AMT)		  =  UNIConvNum(Request("txtVatAmt"),0)
	If Len(Trim(Request("txtVatType"))) > 0 then                                'if VAT is occured(when update)
		iImp_prpaym(C_IMP_VAT_LOC_AMT) = UNIConvNum(Request("txtVatLocAmt"),0)
	Else
		iImp_prpaym(C_IMP_VAT_LOC_AMT) = 0	                                    
	End if	
	iImp_prpaym(C_IMP_PRPAYM_DESC)	  =  Trim(Request("txtPrpaymDesc"))
	If Trim(Request("txtIssuedDt")) <> "" then
		iImp_prpaym(C_IMP_ISSUED_DT)	  =  UNIConvDate(Request("txtIssuedDt"))    '2002.10 patch
	Else
		iImp_prpaym(C_IMP_ISSUED_DT)	  =  UNIConvDate(Request("txtPrpaymDt"))    '2002.10 patch
	End if
	
	iImp_dept(C_IMP_ORG_CHANGE_ID)    =  UCase(Trim(Request("hOrgChangeId")))
	iImp_dept(C_IMP_DEPT_CD)          =  UCase(Trim(Request("txtDeptCd")))
	
'	iImp_b_bank_acct	              =  UCase(Trim(Request("txtBankAcct")))
'	iImp_b_bank			              =  UCase(Trim(Request("txtBankCd")))
	iImp_biz_partner	              =  UCase(Trim(Request("txtBpCd")))
	iImp_b_currency                   =  gCurrency	
	iImp_a_acct_trans_type            = "FP001"
	iImp_a_acct						  =  UCase(Trim(Request("txtAcctCd")))		'2002.10 patch
    iImp_b_tax_biz_area               =  UCase(Trim(Request("txtBizAreaCD")))	'2002.10 patch
    
	'------------------------------------------
	'Com Action Area
	'------------------------------------------

    If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
    End If
    

    iE_prpaym = iPAFG605.F_MANAGE_PRPAYM_SVR(gStrGloBalCollection,iCommandSent,iImp_prpaym,iImp_dept,iImp_b_bank_acct, _
		     iImp_b_bank,iImp_biz_partner,iImp_b_currency,iImp_a_acct_trans_type,iImp_a_acct,iImp_b_tax_biz_area,Trim(Request("txtSpread")),iErrorPosition,I10_a_data_auth)
    
  	if CheckSYSTEMError(Err, True) = True Then					
		Set iPAFG605 = Nothing
		Exit Sub
	End If  

	'------------------------------------------
	'Result data display area
	'------------------------------------------
			
	Response.write " <Script Language=vbscript>" & vbCr
	Response.write " With parent " & vbCr
	Response.write " If .frm1.txtPrpaymNo.Value = """"  Then " & vbCr 
	Response.Write " .frm1.txtPrpaymNo.Value    =  """ & ConvSPChars(iE_prpaym)  & """" & vbCr
	Response.write " End If  " & vbCr                               
	Response.write ".DbSaveOk    " & vbCr                     
	Response.write " End With    " & vbCr
	Response.write " </Script>   " & vbCr

    Set iPAFG605 = Nothing                                                  '☜: Unload Complus
End Sub
	    
Sub SubBizDelete()

	On Error Resume Next
	Err.Clear 

    If Request("txtPrpaymNo") = "" Then										'삭제를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("900002", vbInformation, "", "", I_MKSCRIPT)	    '조회를 먼저 하세요.
		Exit Sub
	End If

	' -- 권한관리추가 
	Const A694_I10_a_data_auth_data_BizAreaCd = 0
	Const A694_I10_a_data_auth_data_internal_cd = 1
	Const A694_I10_a_data_auth_data_sub_internal_cd = 2
	Const A694_I10_a_data_auth_data_auth_usr_id = 3

	Dim I10_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I10_a_data_auth(3)
	I10_a_data_auth(A694_I10_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I10_a_data_auth(A694_I10_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I10_a_data_auth(A694_I10_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I10_a_data_auth(A694_I10_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

    Set iPAFG605 = Server.CreateObject("PAFG605.cFMngPpSvr")

    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If CheckSYSTEMError(Err, True) = True Then					
		Exit Sub
	End If   
    
    '------------------------------------------
    'Data manipulate area
    '------------------------------------------
    Redim iImp_prpaym(21)
	Redim iImp_dept(1)
    
    iCommandSent = "DELETE"
		
	iImp_prpaym(C_IMP_PRPAYM_NO)	  = UCase(Trim(Request("txtPrpaymNo")))
	iImp_prpaym(C_IMP_PRPAYM_DT)	  = UNIConvDate(Request("txtPrpaymDt"))
	iImp_prpaym(C_IMP_PRPAYM_TYPE)	  = Trim(UCase(Request("txtPrpaymType")))
	iImp_prpaym(C_IMP_PAYM_TYPE)	  = UCase(Trim(Request("txtPaymType")))
	iImp_prpaym(C_IMP_REF_NO)         = Trim(Request("txtRefNo"))
	iImp_prpaym(C_IMP_DOC_CUR)        = UCase(Trim(Request("txtDocCur")))
	iImp_prpaym(C_IMP_XCH_RATE)		  = UNIConvNum(Request("txtXchRate"),0)
	iImp_prpaym(C_IMP_NOTE_NO)		  = UCase(Trim(Request("txtNoteNo")))
	iImp_prpaym(C_IMP_PRPAYM_AMT)	  = UNIConvNum(Request("txtPrpaymAmt"),0)
	iImp_prpaym(C_IMP_PRPAYM_LOC_AMT) = UNIConvNum(Request("txtPrpaymLocAmt"),0)
	iImp_prpaym(C_IMP_BAL_AMT)		  = UNIConvNum(Request("txtBalAmt"),0)
	iImp_prpaym(C_IMP_BAL_LOC_AMT)	  = UNIConvNum(Request("txtBalLocAmt"),0)
	iImp_prpaym(C_IMP_GL_NO)          = Trim(Request("txtGlNo"))
	iImp_prpaym(C_IMP_TEMP_GL_NO)     = Trim(Request("txtTempGlNo"))
	iImp_prpaym(C_IMP_PRPAYM_STS)     = ""		
	iImp_prpaym(C_IMP_CONF_FG)        = ""
	iImp_prpaym(C_IMP_PRPAYM_FG)	  = "PP"
	iImp_prpaym(C_IMP_VAT_TYPE)		  = UCase(Trim(Request("txtVatType")))
	iImp_prpaym(C_IMP_VAT_AMT)		  = UNIConvNum(Request("txtVatAmt"),0)
	iImp_prpaym(C_IMP_VAT_LOC_AMT)    = UNIConvNum(Request("txtVatLocAmt"),0)	
	iImp_prpaym(C_IMP_PRPAYM_DESC)	  = Trim(Request("txtPrpaymDesc"))
	iImp_prpaym(C_IMP_ISSUED_DT)	  =  UNIConvDate(Request("txtIssuedDt"))   '2002.10 patch	
    
    iImp_dept(C_IMP_ORG_CHANGE_ID)    = UCase(Trim(Request("hOrgChangeId")))
	iImp_dept(C_IMP_DEPT_CD)          = UCase(Trim(Request("txtDeptCd")))
	
	iImp_b_bank_acct	              = UCase(Trim(Request("txtBankAcct")))
	iImp_b_bank			              = UCase(Trim(Request("txtBankCd")))
	iImp_biz_partner	              = UCase(Trim(Request("txtBpCd")))
	iImp_b_currency                   = gCurrency	
	iImp_a_acct_trans_type            = "FP001"
	iImp_a_acct						  =  UCase(Trim(Request("txtAcctCd")))		'2002.10 patch
    iImp_b_tax_biz_area               =  UCase(Trim(Request("txtBizAreaCD")))	'2002.10 patch    
    '------------------------------------------
    'Com Action Area
    '------------------------------------------
	Call iPAFG605.F_MANAGE_PRPAYM_SVR(gStrGloBalCollection,iCommandSent,iImp_prpaym,iImp_dept,iImp_b_bank_acct, _
		  iImp_b_bank,iImp_biz_partner,iImp_b_currency,iImp_a_acct_trans_type,iImp_a_acct,iImp_b_tax_biz_area,,,I10_a_data_auth)
        
    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
  	if CheckSYSTEMError(Err, True) = True Then					
		Set iPAFG605 = Nothing
		Exit Sub
	End If  

	'------------------------------------------
	'Result data display area
	'------------------------------------------

	Response.Write " <Script Language=vbscript>" & vbCr
	Response.Write " Call parent.DbDeleteOk() " & vbCr
	Response.Write " </Script> "			     & vbCr
	        
    Set iPAFG605 = Nothing                                                   '☜: Unload Complus

End sub	
%>
