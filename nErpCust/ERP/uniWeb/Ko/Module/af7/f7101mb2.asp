<%'======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : Prereceipt
'*  3. Program ID           : f7101mb2
'*  4. Program Name         : 선수금 등록-저장 
'*  5. Program Desc         : 선수금 등록 
'*  6. Modified date(First) : 2000/09/25
'*  7. Modified date(Last)  : 2002/11/18
'*  8. Modifier (First)     : 조익성 
'*  9. Modifier (Last)      : Jeong Yong Kyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

														'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True														'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%													                        '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜: 
Err.Clear 

Call LoadBasisGlobalInf()

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim iPAFG705															    '☆ : 저장용 ComPlus Dll 사용 변수 
Dim IntRows
Dim IntCols
Dim lgIntFlgMode
Dim LngMaxRow
dim iPrrcpt_No
Dim iArrPrrcpt
Dim iArrDept
dim iStrBizPartner
dim iStrCurrency
Dim iStrTransType
Dim iCommandSent
Dim istrTaxBizArea
Dim iErrorPosition

Const C_PRRCPT_NO = 0
Const C_PRRCPT_DT = 1
Const C_REF_NO = 2
Const C_DOC_CUR = 3
Const C_XCH_RATE = 4
Const C_PRRCPT_AMT = 5
Const C_LOC_PRRCPT_AMT = 6
Const C_PRRCPT_STS = 7
Const C_CONF_FG = 8
Const C_PRRCPT_FG = 9
Const C_PRRCPT_DESC = 10
Const C_PRRCPT_TYPE = 11
Const C_VAT_TYPE = 12
Const C_VAT_AMT = 13
Const C_VAT_LOC_AMT = 14
Const C_ISSUED_DT = 15
Const C_PROJECT_NO = 16
Const C_LIMIT_FG = 17

Const C_CHANGEORGID = 0
Const C_DEPT_CD = 1

	' -- 저장용 
	' -- 권한관리추가 
	Const A750_I1_a_data_auth_data_BizAreaCd = 0
	Const A750_I1_a_data_auth_data_internal_cd = 1
	Const A750_I1_a_data_auth_data_sub_internal_cd = 2
	Const A750_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

	Set iPAFG705 = Server.CreateObject("PAFG705.cFMngPrSvr")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then					
	   Response.End 
	End If    

	'-----------------------
	'Data manipulate area
	'-----------------------														'⊙: Single 데이타 저장 
	Redim iArrPrrcpt(C_LIMIT_FG)
	Redim iArrDept(1)

	iStrTransType	             = "FR001"
	iArrPrrcpt(C_PRRCPT_NO)      = Trim(Request("txtPrrcptNo"))
	iArrPrrcpt(C_PRRCPT_DT)      = UNIConvDate(Request("txtPrrcptDt"))
	iArrPrrcpt(C_REF_NO)         = Trim(Request("txtRefNo"))
	iArrPrrcpt(C_DOC_CUR)        = UCase(Trim(Request("txtDocCur")))
	iArrPrrcpt(C_XCH_RATE)       = UNIConvNum(Request("txtXchRate"),0)
	iArrPrrcpt(C_PRRCPT_AMT)     = UNIConvNum(Request("txtPrrcptAmt"),0)
	iArrPrrcpt(C_LOC_PRRCPT_AMT) = ""
	iArrPrrcpt(C_PRRCPT_STS)     = ""
	iArrPrrcpt(C_CONF_FG)        = ""
	iArrPrrcpt(C_PRRCPT_FG)      = "PC"
	iArrPrrcpt(C_PRRCPT_DESC)    = Trim(Request("txtPrrcptDesc"))
	iArrPrrcpt(C_PRRCPT_TYPE)    = UCase(Trim(Request("txtPrrcptType")))
	iArrPrrcpt(C_VAT_TYPE)       = UCase(Trim(Request("txtVatType")))
	iArrPrrcpt(C_VAT_AMT)        = UNIConvNum(Request("txtVatAmt"),0)
	iArrPrrcpt(C_VAT_LOC_AMT)    = UNIConvNum(Request("txtVatLocAmt"),0)
	iArrPrrcpt(C_PROJECT_NO)	 = UCase(Trim(Request("txtProjectNo")))
	iArrPrrcpt(C_LIMIT_FG)		 = Trim(Request("txtLimitFg"))
		
	If Len(Trim(Request("txtVatType"))) > 0 then                                'if VAT is occured(when update)
		iArrPrrcpt(C_VAT_LOC_AMT) = UNIConvNum(Request("txtVatLocAmt"),0)
	Else
		iArrPrrcpt(C_VAT_LOC_AMT) = 0	                                    
	End If		
	If Trim(Request("txtIssuedDt")) = "" Or isnull(Trim(Request("txtIssuedDt"))) Then
		iArrPrrcpt(C_ISSUED_DT)  = UNIConvDate(Request("txtPrrcptDt"))
	Else
		iArrPrrcpt(C_ISSUED_DT)  = UNIConvDate(Request("txtIssuedDt"))
	End if		

	iStrBizPartner               = UCase(Trim(Request("txtBpCd")))

	iArrdept(C_DEPT_CD)          = UCase(Trim(Request("txtDeptCd")))
	iArrDept(C_CHANGEORGID)      = Trim(request("hOrgChangeId"))

	iStrCurrency                 = gCurrency
	iStrTaxBizArea               = UCase(Trim(Request("txtBizAreaCD")))
	

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	iPrrcptNo = iPAFG705.F_MANAGE_PRRCPT_SVR(gStrGloBalCollection,iCommandSent,iStrTransType, _
	                                         iStrCurrency,iArrDept,iStrBizPartner,iArrPrrcpt, _
	                                         iStrTaxBizArea,Trim(Request("txtSpread")),iErrorPosition,I1_a_data_auth)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then					
		Set iPAFG705 = Nothing
		Response.End 
	End If    
	
'    if CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then					
'	    Set iPAFG705 = Nothing
'	    Response.End 
'	End If    

	Response.Write " <Script Language=vbscript>                                     " & vbCr
	Response.Write " With parent                                                    " & vbCr   '☜: 화면 처리 ASP 를 지칭함 
	Response.Write " .frm1.txtPrrcptNo.Value	= """ & ConvSPChars(iPrrcptNo) & """" & vbCr
	Response.Write " .DbSaveOk                                                      " & vbCr
	Response.Write "  End With                                                      " & vbCr
	Response.Write " </Script>                                                      " & vbCr

	Set iPAFG705 = Nothing															           '☜: Unload Complus
%>
