<%'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : Prereceipt
'*  3. Program ID           : f7101mb1
'*  4. Program Name         : 선수금정보 등록 
'*  5. Program Desc         : 선수금정보 등록의 조회로직 
'*  6. Modified date(First) : 2000/09/25
'*  7. Modified date(Last)  : 2002/11/18
'*  8. Modifier (First)     : 조익성 
'*  9. Modifier (Last)      : Jeong Yogn Kyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
Option Explicit
														'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True														'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													

On Error Resume Next	
Err.Clear 																	'☜: 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

Call HideStatusWnd

Dim iPAFG705																'☆ : 조회용 ComPlus Dll 사용 변수 
Dim StrNextKey																' 다음 값 
Dim lgStrPrevKey															' 이전 값 
Dim LngMaxRow																' 현재 그리드의 최대Row
Dim LngRow
Dim iCommandSent
Dim iE_arrPrrcpt
Dim iEG_Spread     
Dim iArrPrrcpt
Dim strData
Dim lgCurrency

Dim LookUp_states_char 

Const C_PRRCPT_NO = 0
Const C_PRRCPT_TYPE = 1
Const C_PRRCPT_TYPE_NM = 2
Const C_PRRCPT_DT = 3
Const C_PRRCPT_DEPT_CD = 4
Const C_PRRCPT_DEPT_NM = 5
Const C_PRRCPT_BP_CD = 6
Const C_PRRCPT_BP_NM = 7
Const C_PRRCPT_REF_NO = 8
Const C_PRRCPT_DOC_CUR = 9
Const C_PRRCPT_XCH_RATE = 10
Const C_PRRCPT_PRRCPT_AMT = 11
Const C_PRRCPT_LOC_PRRCPT_AMT = 12
Const C_PRRCPT_CLS_AMT = 13
Const C_PRRCPT_LOC_CLS_AMT = 14
Const C_PRRCPT_STTL_AMT = 15
Const C_PRRCPT_LOC_STTL_AMT = 16
Const C_PRRCPT_BAL_AMT = 17
Const C_PRRCPT_LOC_BAL_AMT = 18
Const C_PRRCPT_VAT_TYPE = 19
Const C_PRRCPT_VAT_TYPE_NM = 20
Const C_PRRCPT_VAT_AMT = 21
Const C_PRRCPT_VAT_LOC_AMT = 22
Const C_PRRCPT_GL_NO = 23
Const C_PRRCPT_TEMP_GL_NO = 24
Const C_PRRCPT_PRRCPT_DESC = 25
Const C_PRRCPT_IO_FG = 26
Const C_PRRCPT_IO_FG_NM = 27	
Const C_PRRCPT_REPORT_BIZ_AREA_CD = 28
Const C_PRRCPT_REPORT_BIZ_AREA_NM = 29
Const C_PRRCPT_ISSUED_DT = 30
Const C_PRRCPT_PROJECT_NO = 31
Const C_PRRCPT_LIMIT_FG = 32	

Const C_SPREAD_SEQ = 0
Const C_SPREAD_RCPT_TYPE =1
Const C_SPREAD_RCPT_TYPE_NM = 2
Const C_SPREAD_AMT = 3
Const C_SPREAD_LOC_AMT = 4
Const C_SPREAD_BANK_CD = 5
Const C_SPREAD_BANK_NM = 6 
Const C_SPREAD_BANK_ACCT_NO = 7
Const C_SPREAD_NOTE_NO = 8
Const C_SPREAD_ACCT_CD = 9
Const C_SPREAD_ACCT_NM = 10	
Const C_SPREAD_C_STTL_DESC = 11	

Const C_LOOKUP_NO = 0
Const C_LOOKUP_FG = 1

	' -- 조회용 
	' -- 권한관리추가 
	Const A750_I1_a_data_auth_data_BizAreaCd = 0
	Const A750_I1_a_data_auth_data_internal_cd = 1
	Const A750_I1_a_data_auth_data_sub_internal_cd = 2
	Const A750_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
		
	If Request("lgStrPrevKey") = "" then
		lgStrPrevKey = 0
	Else
		lgStrPrevKey = Request("lgStrPrevKey")
	End if
	    
	Set iPAFG705 = Server.CreateObject("PAFG705.cFLkUpPrSvr")

	If CheckSYSTEMError(Err, True) = True Then					
		Response.End 
	End If   

	Redim iArrPrrcpt(1)
		
	iCommandSent = Request("txtCommand")
	    
	iArrPrrcpt(C_LOOKUP_NO) = UCase(Trim(Request("txtPrrcptNo")))
	iArrPrrcpt(C_LOOKUP_FG) = "PC"
	
	LookUp_states_char = iPAFG705.F_LOOKUP_PRRCPT_SVR(gStrGloBalCollection,iCommandSent,iArrPrrcpt,iE_arrPrrcpt,iEG_Spread,I1_a_data_auth)

	If CheckSYSTEMError(Err, True) = True Then					
		Set iPAFG705 = Nothing
		Response.End 
	End If  

	LngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
	lgCurrency = ConvSPChars(iE_arrPrrcpt(C_PRRCPT_DOC_CUR))

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent " & vbCr
	Response.Write ".frm1.txtPrrcptNo.value		= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_NO))         & """" & vbCr
	Response.Write ".frm1.txtPrrcptType.value	= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_TYPE))       & """" & vbCr
	Response.Write ".frm1.txtPrrcptTypeNm.value	= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_TYPE_NM))    & """" & vbCr
	If ConvSPChars(iE_arrPrrcpt(C_PRRCPT_LIMIT_FG)) = "Y" Then
		Response.Write ".frm1.chkLimitFg.checked 	= True  " & vbcr
		Response.Write ".frm1.txtLimitFg.value		= ""Y"" " & vbcr
	Else
		Response.Write ".frm1.chkLimitFg.checked	= False " & vbcr
		Response.Write ".frm1.txtLimitFg.value		= ""N"" " & vbcr
	End If	
	Response.Write ".frm1.txtPrrcptDt.Text		= """ & UNIDateClientFormat(iE_arrPrrcpt(C_PRRCPT_DT)) & """" & vbCr
	Response.Write ".frm1.txtDeptCd.value		= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_DEPT_CD))    & """" & vbCr
	Response.Write ".frm1.txtDeptNm.value		= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_DEPT_NM))    & """" & vbCr
	Response.Write ".frm1.txtBpCd.value			= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_BP_CD))      & """" & vbCr
	Response.Write ".frm1.txtBpNm.value			= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_BP_NM))      & """" & vbCr
	Response.Write ".frm1.txtRefNo.value		= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_REF_NO))     & """" & vbCr
	Response.Write ".frm1.txtDocCur.value		= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_DOC_CUR))	   & """" & vbCr
	Response.Write ".frm1.txtProjectNo.value	= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_PROJECT_NO))	   & """" & vbCr
    Response.Write ".frm1.txtPrrcptAmt.Text	    = """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_PRRCPT_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")               & """" & vbCr	
	Response.Write ".frm1.txtClsLocAmt.Text	    = """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_LOC_CLS_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")    & """" & vbCr	
	Response.Write ".frm1.txtSttlLocAmt.Text	= """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_LOC_STTL_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")   & """" & vbCr
	Response.Write ".frm1.txtBalLocAmt.Text	    = """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_LOC_BAL_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")    & """" & vbCr
	Response.Write ".frm1.txtVatAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_VAT_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                  & """" & vbCr

	If gIsShowLocal <> "N" Then	
	
		Response.Write ".frm1.txtXchRate.Text		= """ & UNINumClientFormat(iE_arrPrrcpt(C_PRRCPT_XCH_RATE), ggExchRate.DecPoint, 0)                                            & """" & vbCr
		Response.Write ".frm1.txtPrrcptLocAmt.Text	= """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_LOC_PRRCPT_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
		Response.Write ".frm1.txtClsAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_CLS_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                  & """" & vbCr
		Response.Write ".frm1.txtSttlAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_STTL_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                 & """" & vbCr
		Response.Write ".frm1.txtBalAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_BAL_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                  & """" & vbCr
		Response.Write ".frm1.txtVatLocAmt.Text	    = """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_VAT_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")    & """" & vbCr
	Else

		Response.Write ".frm1.txtXchRate.Value		= """ & UNINumClientFormat(iE_arrPrrcpt(C_PRRCPT_XCH_RATE), ggExchRate.DecPoint, 0)                                            & """" & vbCr
		Response.Write ".frm1.txtPrrcptLocAmt.Value	= """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_LOC_PRRCPT_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
		Response.Write ".frm1.txtClsAmt.Value		= """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_CLS_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                  & """" & vbCr
		Response.Write ".frm1.txtSttlAmt.Value		= """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_STTL_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                 & """" & vbCr
		Response.Write ".frm1.txtBalAmt.Value		= """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_BAL_AMT), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                  & """" & vbCr
		Response.Write ".frm1.txtVatLocAmt.Value    = """ & UNIConvNumDBToCompanyByCurrency(iE_arrPrrcpt(C_PRRCPT_VAT_LOC_AMT), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")    & """" & vbCr
	End if

	Response.Write ".frm1.txtVatType.value		= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_VAT_TYPE))           & """" & vbCr
	Response.Write ".frm1.txtVatTypeNm.value    = """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_VAT_TYPE_NM))        & """" & vbCr

'	If Len(Trim(iE_arrPrrcpt(C_PRRCPT_VAT_TYPE))) > 0 then
		Response.Write ".frm1.txtBizAreaCD.value	= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_REPORT_BIZ_AREA_CD)) & """" & vbCr
		Response.Write ".frm1.txtBizAreaNM.value	= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_REPORT_BIZ_AREA_NM)) & """" &	vbCr
		Response.Write ".frm1.txtIssuedDt.text		= """ & UNIDateClientFormat(iE_arrPrrcpt(C_PRRCPT_ISSUED_DT))  & """" & vbCr
'	Else
'		Response.Write ".frm1.txtBizAreaCD.value	= """"" & vbCr 
'		Response.Write ".frm1.txtBizAreaNM.value	= """"" & VbCr				
'		Response.Write ".frm1.txtIssuedDt.text		= """"" & vbCr									
'	End if					
			
	Response.Write ".frm1.txtTempGlNo.value		= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_TEMP_GL_NO))         & """" & vbCr
	Response.Write ".frm1.txtGlNo.value			= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_GL_NO))              & """" & vbCr
	Response.Write ".frm1.txtPrrcptDesc.value	= """ & ConvSPChars(iE_arrPrrcpt(C_PRRCPT_PRRCPT_DESC))        & """" & vbCr
	Response.Write ".frm1.txtAfterLookUp.value	= """ & ConvSPChars(LookUp_states_char)                        & """" & vbCr
		
	LngMaxRow = .frm1.vspdData.MaxRows										                            'Save previous Maxrow

	For LngRow = 0 To UBOUND(iEG_Spread,1)
	    strData = strData & Chr(11) & iEG_Spread(LngRow,C_SPREAD_SEQ)
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_RCPT_TYPE))
		strData = strData & Chr(11) & " "    'popup button
		strData = strData & Chr(11) & ConvSPChars(iEG_Spread(LngRow,C_SPREAD_RCPT_TYPE_NM))
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
'		strData = strData & Chr(11) & " "    'popup button
		strData = strData & Chr(11) & LngMaxRow + LngRow + 1
	    strData = strData & Chr(11) & Chr(12)
	Next
	Response.Write ".ggoSpread.Source = .frm1.vspdData "          & vbCr
	Response.Write ".ggoSpread.SSShowData """ & strData    & """" & vbCr
	Response.Write ".lgStrPrevKey = """ & ConvSPChars(StrNextKey) & """" & vbCr
	Response.Write ".DbQueryOk "                                         & vbCr
	Response.Write " End With  "                                         & vbCr
	Response.Write " </Script> "                                         & vbCr

	Set iPAFG705 = Nothing
%>
