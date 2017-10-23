<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3101mb1
'*  4. Program Name         : 예적금등록 
'*  5. Program Desc         : Register of Deposit Master
'*  6. Comproxy List        : FD0011, FD0019
'*  7. Modified date(First) : 2000.09.19
'*  8. Modified date(Last)  : 2002.06.20
'*  9. Modifier (First)     : Kim, Jong Hwan
'* 10. Modifier (Last)      : JANG YOON KI
'* 11. Comment              :
'=======================================================================================================


On Error Resume Next                                                            '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status


Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")

Call HideStatusWnd


Dim lgOpModeCRUD
Dim txtBankCd, txtBankAcctNo
'---------------------------------------Common-----------------------------------------------------------
                                                       '☜: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
txtBankCd         = Trim(Request("txtBankCd"))
txtBankAcctNo	  = Trim(Request("txtBankAcctNo"))	
'------ Developer Coding part (Start ) ------------------------------------------------------------------

' 권한관리 추가 
Dim lgAuthBizAreaCd			' 사업장 
Dim lgInternalCd			' 내부부서 
Dim lgSubInternalCd			' 내부부서(하위포함)
Dim lgAuthUsrID				' 개인 

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
		Call SubBizQuery()
    Case CStr(UID_M0002) 
        Call SubBizSave()
    Case CStr(UID_M0003)                                                         '☜: Delete
        Call SubBizDelete()
End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

	Const C_BANK_CD = 0
	Const C_BANK_NM = 1
	Const C_BANK_ACCT_NO = 2
	Const C_DPST_FG = 3
	Const C_DPST_TYPE = 4
	Const C_DEPT_CD = 5
	Const C_DEPT_NM = 6
	Const C_START_DT = 7
	Const C_TRANS_STS = 8
	Const C_BANK_ACCT_FG = 9
	Const C_BANK_RATE = 10
	Const C_DPST_NM = 11
	Const C_DOC_CUR = 12
	Const C_XCH_RATE = 13
	Const C_AMT = 14
	Const C_LOC_AMT = 15
	Const C_END_DT = 16
	Const C_PAYM_DT = 17		
	Const C_PAYM_PERIOD = 18
	Const C_PAYM_CNT = 19
	Const C_TOT_PAYM_CNT = 20
	Const C_PAYM_AMT = 21
	Const C_PAYM_LOC_AMT = 22
	Const C_CONTRACT_AMT = 23
	Const C_CONTRACT_LOC_AMT = 24
	Const C_CNCL_DT = 25
	Const C_CNCL_XCH_RATE = 26
	Const C_CNCL_INT_RATE = 27
	Const C_CNCL_CAPITAL_AMT = 28
	Const C_CNCL_CAP_LOC_AMT = 29
	Const C_CNCL_INT_AMT = 30
	Const C_CNCL_INT_LOC_AMT = 31
	Const C_CNCL_AMT = 32
	Const C_CNCL_LOC_AMT = 33
	Const C_DPST_DESC = 34

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear       
	 
	Dim PAFG305LIST
	Dim E5_f_dpst

	Redim E5_f_dpst(34)

	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

    Const A817_I1_a_data_auth_data_BizAreaCd = 0
    Const A817_I1_a_data_auth_data_internal_cd = 1
    Const A817_I1_a_data_auth_data_sub_internal_cd = 2
    Const A817_I1_a_data_auth_data_auth_usr_id = 3 
    
  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A817_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I1_a_data_auth(A817_I1_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I1_a_data_auth(A817_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I1_a_data_auth(A817_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	
	If txtBankCd = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)	'조회 조건값이 비어있습니다.
		Response.End 
	End If

	If txtBankAcctNo = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)	'조회 조건값이 비어있습니다.
		Response.End 
	End If

	Set PAFG305LIST = server.CreateObject ("PAFG305.cFLkUpDpstSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
	
    Call PAFG305LIST.FD0019_LOOKUP_DPST_SVR(gStrGlobalCollection,txtBankCd,txtBankAcctNo,E5_f_dpst, I1_a_data_auth)
    
    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG305LIST = nothing		
		Exit Sub
    End If
    														
    Set PAFG305LIST = nothing
	
		Response.Write " <Script Language=vbscript>															 " & vbCr
		Response.Write " With parent.frm1																	 " & vbCr
		Response.Write "	.txtBankCd.Value		=	""" & ConvSPChars(E5_f_dpst(C_BANK_CD))			& """" & vbCr
		Response.Write "	.txtBankNm.Value		=	""" & ConvSPChars(E5_f_dpst(C_BANK_NM))			& """" & vbCr
		Response.Write "	.txtBankAcctNo.Value	=	""" & ConvSPChars(E5_f_dpst(C_BANK_ACCT_NO))	& """" & vbCr	
		Response.Write "	.cboDpstFg.Value		=	""" & ConvSPChars(E5_f_dpst(C_DPST_FG))			& """" & vbCr
		Response.Write "	.cboDpstType.Value		=	""" & ConvSPChars(E5_f_dpst(C_DPST_TYPE))		& """" & vbCr	
		Response.Write "	.txtStartDt.Text		=	""" & UNIDateClientFormat(E5_f_dpst(C_START_DT)) & """" & vbCr
		Response.Write "	.txtDeptCD.Value		=	""" & ConvSPChars(E5_f_dpst(C_DEPT_CD))			& """" & vbCr
		Response.Write "	.txtDeptNm.Value		=	""" & ConvSPChars(E5_f_dpst(C_DEPT_NM))			& """" & vbCr
		Response.Write "	.cboTransSts.Value		=	""" & ConvSPChars(E5_f_dpst(C_TRANS_STS))		& """" & vbCr
		Response.Write "	.cboBankAcctFg.Value	=	""" & ConvSPChars(E5_f_dpst(C_BANK_ACCT_FG))	& """" & vbCr
		Response.Write "	.txtBankRate.Text		=	""" & ConvSPChars(E5_f_dpst(C_BANK_RATE))		& """" & vbCr	
		Response.Write "	.txtDpstNm.Value		=	""" & ConvSPChars(E5_f_dpst(C_DPST_NM))			& """" & vbCr
		Response.Write "	.txtDocCur.Value		=	""" & ConvSPChars(E5_f_dpst(C_DOC_CUR))			& """" & vbCr
		Response.Write "	.txtXchRate.Text		=	""" & UNINumClientFormat(E5_f_dpst(C_XCH_RATE),			ggExchRate.DecPoint	,0)			& """" & vbCr
		Response.Write "	.txtAmt.Text			=	""" & UNINumClientFormat(E5_f_dpst(C_AMT),				ggAmtOfMoney.DecPoint	,0)				& """" & vbCr
		Response.Write "	.txtLocAmt.Text		=	""" & UNINumClientFormat(E5_f_dpst(C_LOC_AMT),			ggAmtOfMoney.DecPoint	,0)			& """" & vbCr
		Response.Write "	.txtEndDt.Text			=	""" & UNIDateClientFormat(E5_f_dpst(C_END_DT))	& """" & vbCr
		Response.Write "	.txtPaymDt.Text		=	""" & E5_f_dpst(C_PAYM_DT)	& """" & vbCr
		Response.Write "	.txtPaymPeriod.Value	=	""" & E5_f_dpst(C_PAYM_PERIOD)		& """" & vbCr
		Response.Write "	.txtPaymCnt.Value		=	""" & E5_f_dpst(C_PAYM_CNT)			& """" & vbCr
		Response.Write "	.txtTotPaymCnt.Value	=	""" & E5_f_dpst(C_TOT_PAYM_CNT)		& """" & vbCr
		Response.Write "	.txtPaymLocAmt.Text	=	""" & UNINumClientFormat(E5_f_dpst(C_PAYM_LOC_AMT),		ggAmtOfMoney.DecPoint	,0)		& """" & vbCr
		Response.Write "	.txtPaymAmt.Text		=	""" & UNINumClientFormat(E5_f_dpst(C_PAYM_AMT),			ggAmtOfMoney.DecPoint	,0)			& """" & vbCr	
		Response.Write "	.txtContractAmt.Text	=	""" & UNINumClientFormat(E5_f_dpst(C_CONTRACT_AMT),		ggAmtOfMoney.DecPoint	,0)		& """" & vbCr
		Response.Write "	.txtContractLocAmt.Text=	""" & UNINumClientFormat(E5_f_dpst(C_CONTRACT_LOC_AMT),	ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
		Response.Write "	.txtCnclDt.Text		=	""" & UNIDateClientFormat(E5_f_dpst(C_CNCL_DT))	& """" & vbCr
		Response.Write "	.txtCnclXchRate.Text	=	""" & UNINumClientFormat(E5_f_dpst(C_CNCL_XCH_RATE),	ggExchRate.DecPoint	,0)	& """" & vbCr
		Response.Write "	.txtCnclIntRate.Text	=	""" & UNINumClientFormat(E5_f_dpst(C_CNCL_INT_RATE),	ggExchRate.DecPoint	,0)	& """" & vbCr
		Response.Write "	.txtCnclCapitalAmt.Text=	""" & UNINumClientFormat(E5_f_dpst(C_CNCL_CAPITAL_AMT),	ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
		Response.Write "	.txtCnclCapLocAmt.Text	=	""" & UNINumClientFormat(E5_f_dpst(C_CNCL_CAP_LOC_AMT),	ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
		Response.Write "	.txtCnclIntAmt.Text	=	""" & UNINumClientFormat(E5_f_dpst(C_CNCL_INT_AMT),		ggAmtOfMoney.DecPoint	,0)		& """" & vbCr
		Response.Write "	.txtCnclIntLocAmt.Text	=	""" & UNINumClientFormat(E5_f_dpst(C_CNCL_INT_LOC_AMT),	ggAmtOfMoney.DecPoint	,0)	& """" & vbCr
		Response.Write "	.txtCnclAmt.Text		=	""" & UNINumClientFormat(E5_f_dpst(C_CNCL_AMT),			ggAmtOfMoney.DecPoint	,0)			& """" & vbCr
		Response.Write "	.txtCnclLocAmt.Text	=	""" & UNINumClientFormat(E5_f_dpst(C_CNCL_LOC_AMT),		ggAmtOfMoney.DecPoint	,0)		& """" & vbCr
		Response.Write "	.txtDpstDesc.Value		=	""" & ConvSPChars(E5_f_dpst(C_DPST_DESC))   & """" & vbCr
		Response.Write "	parent.DbQueryOk													 " & vbCr
		Response.Write "End With																 " & vbCr
		Response.Write "</Script>																 " & vbCr
End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data To Db
'============================================================================================================
Sub SubBizSave()
	Dim I1_b_bank
	Dim I2_b_bank_acct
	Dim I3_b_acct_dept
	Dim iarrData
	Dim I5_f_dpst
	Dim PAFG305CU
	Dim lgIntFlgMode
		
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear        

	Const C_BANK_CD = 0
	Const C_BANK_NM = 1

	Redim I1_b_bank(C_BANK_NM) 
		I1_b_bank(C_BANK_CD) = Trim(Request("txtBankCD")) 
		I1_b_bank(C_BANK_NM) = Trim(Request("txtBankNM")) 


		I2_b_bank_acct = Trim(Request("txtBankAcctNo"))
		
	Const C_CHG_ORG_ID = 0
	Const C_DEPT_CD = 1

	ReDim I3_b_acct_dept(C_DEPT_CD)
		I3_b_acct_dept(C_CHG_ORG_ID) = Trim(request("horgchangeid"))
		I3_b_acct_dept(C_DEPT_CD) = Trim(Request("txtDeptCD"))

	Const C_DPST_FG = 0
	Const C_DPST_TYPE = 1
	Const C_BANK_ACCT_FG = 2
	Const C_LOAN_TYPE = 3
	Const C_BANK_RATE = 4
	Const C_START_DT = 5
	Const C_END_DT = 6
	Const C_DOC_CUR = 7
	Const C_XCH_RATE = 8
	Const C_TRANS_STS = 9
	Const C_CONTRACT_AMT = 10
	Const C_CONTRACT_LOC_AMT = 11
	Const C_PAYM_DT = 12
	Const C_PAYM_PERIOD = 13
	Const C_PAYM_CNT = 14
	Const C_TOT_PAYM_CNT = 15
	Const C_PAYM_AMT = 16
	Const C_PAYM_LOC_AMT = 17
	Const C_CNCL_DT = 18
	Const C_CNCL_XCH_RATE = 19
	Const C_CNCL_CAPITAL_AMT = 20
	Const C_CNCL_CAP_LOC_AMT = 21
	Const C_CNCL_INT_RATE = 22
	Const C_CNCL_INT_AMT = 23
	Const C_CNCL_INT_LOC_AMT = 24
	Const C_CNCL_AMT = 25
	Const C_CNCL_LOC_AMT = 26
	Const C_DPST_NM = 27
	Const C_DPST_DESC = 28

	ReDim iarrData(C_DPST_DESC)    
			
	iarrData(C_DPST_FG)				= Request("cboDpstFg")							'예적금구분				        
	iarrData(C_DPST_TYPE)			= Request("cboDpstType")						'예적금유형 
	iarrData(C_BANK_ACCT_FG)		= Request("cboBankAcctFg")						'계좌유형 
	iarrData(C_LOAN_TYPE)			= ""			        
	iarrData(C_BANK_RATE)			= UNIConvNum(Request("txtBankRate"), 0)			'이율		        
	iarrData(C_START_DT)			= UniConvDate(Request("txtStartDt"))			'거래시작일		        
	iarrData(C_END_DT)				= UniConvDate(Request("txtEndDt"))				'만기일 
	iarrData(C_DOC_CUR)				= Trim(Request("txtDocCur"))					'거래통화 
	iarrData(C_XCH_RATE)			= UNIConvNum(Request("txtXchRate"), 0)			'환율 
	iarrData(C_TRANS_STS)			= Request("cboTransSts")						'거래상태 
	iarrData(C_CONTRACT_AMT)		= UNIConvNum(Request("txtContractAmt"), 0)		'계약금액 
	iarrData(C_CONTRACT_LOC_AMT)	= UNIConvNum(Request("txtContractLocAmt"), 0)	'계약금액(자국)
	iarrData(C_PAYM_DT)				= UNIConvNum(Request("txtPaymDt"),0)			'납입일 
	iarrData(C_PAYM_PERIOD)			= UNIConvNum(Request("txtPaymPeriod"), 0)		'납입주기 
	iarrData(C_PAYM_CNT)			= UNIConvNum(Request("txtPaymCnt"), 0)			'불입회수 
	iarrData(C_TOT_PAYM_CNT)		= UNIConvNum(Request("txtTotPaymCnt"), 0)		'총불입회수 
	iarrData(C_PAYM_AMT)			= UNIConvNum(Request("txtPaymAmt"), 0)			'월납입액 
	iarrData(C_PAYM_LOC_AMT)		= UNIConvNum(Request("txtPaymLocAmt"), 0)		'월납입액(자국)
	iarrData(C_CNCL_DT)				= UNIConvDate(Request("txtCnclDt"))				'해약일자     
	iarrData(C_CNCL_XCH_RATE)		= UNIConvNum(Request("txtCnclXchRate"), 0)		'해약시환율 
	iarrData(C_CNCL_CAPITAL_AMT)	= UNIConvNum(Request("txtCnclCapitalAmt"), 0)	'해약시원금 
	iarrData(C_CNCL_CAP_LOC_AMT)	= UNIConvNum(Request("txtCnclCapLocAmt"), 0)	'해약시원금(자국)
	iarrData(C_CNCL_INT_RATE)		= UNIConvNum(Request("txtCnclIntRate"), 0)		'해약시이자율 
	iarrData(C_CNCL_INT_AMT)		= UNIConvNum(Request("txtCnclIntAmt"), 0)		'해약시이자 
	iarrData(C_CNCL_INT_LOC_AMT)	= UNIConvNum(Request("txtCnclIntLocAmt"), 0)	'해약시이자(자국)
	iarrData(C_CNCL_AMT)			= UNIConvNum(Request("txtCnclAmt"), 0)			'해약금액 
	iarrData(C_CNCL_LOC_AMT)		= UNIConvNum(Request("txtCnclLocAmt"), 0)		'해약금액(자국)
	iarrData(C_DPST_NM)				= Request("txtDpstNm")	'가입사유 
	iarrData(C_DPST_DESC)			= Request("txtDpstDesc")	'적요 

	
	I5_f_dpst = gCurrency
	
	Dim  I6_a_data_auth
	' -- 권한관리 
    Const A816_I4_a_data_auth_data_BizAreaCd = 0
    Const A816_I4_a_data_auth_data_internal_cd = 1
    Const A816_I4_a_data_auth_data_sub_internal_cd = 2
    Const A816_I4_a_data_auth_data_auth_usr_id = 3 
    
  	Redim I6_a_data_auth(3)
	I6_a_data_auth(A816_I4_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I6_a_data_auth(A816_I4_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I6_a_data_auth(A816_I4_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I6_a_data_auth(A816_I4_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
		
    Set PAFG305CU = server.CreateObject("PAFG305.cFMngDpstSvr")   

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
     
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
    
    Select Case lgIntFlgMode
		Case  OPMD_CMODE                                                             '☜ : Create
			Call PAFG305CU.FD0011_MANAGE_DPST_SVR(gStrGlobalCollection,"CREATE",I1_b_bank,I2_b_bank_acct,I3_b_acct_dept,iarrData,I5_f_dpst, I6_a_data_auth)
        Case  OPMD_UMODE           
			Call PAFG305CU.FD0011_MANAGE_DPST_SVR(gStrGlobalCollection,"UPDATE",I1_b_bank,I2_b_bank_acct,I3_b_acct_dept,iarrData,I5_f_dpst, I6_a_data_auth)
    End Select

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG305CU = nothing
		Exit Sub	
    End If

    Set PAFG305CU = nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr    
	
End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : DELETE Data 
'============================================================================================================

Sub SubBizDelete()
	Dim PAFG305D
	Dim I1_b_bank
	Dim I2_b_bank_acct

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear        

	Dim I6_a_data_auth
	' -- 권한관리 
    Const A816_I4_a_data_auth_data_BizAreaCd = 0
    Const A816_I4_a_data_auth_data_internal_cd = 1
    Const A816_I4_a_data_auth_data_sub_internal_cd = 2
    Const A816_I4_a_data_auth_data_auth_usr_id = 3 
    
  	Redim I6_a_data_auth(3)
	I6_a_data_auth(A816_I4_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I6_a_data_auth(A816_I4_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I6_a_data_auth(A816_I4_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I6_a_data_auth(A816_I4_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
			
	Const C_BANK_CD = 0
	Const C_BANK_NM = 1

	Redim I1_b_bank(C_BANK_NM) 
	I1_b_bank(C_BANK_CD) = Trim(Request("txtBankCD")) 
	I1_b_bank(C_BANK_NM) = Trim(Request("txtBankNM")) 

	I2_b_bank_acct = Trim(Request("txtBankAcctNo"))
	
    Set PAFG305D = server.CreateObject ("PAFG305.cFMngDpstSvr")    
    
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
	
	
    Call PAFG305D.FD0011_MANAGE_DPST_SVR(gStrGlobalCollection,"DELETE",I1_b_bank,I2_b_bank_acct,,,, I6_a_data_auth)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG305D = nothing
		Exit Sub
    End If
	 
    Set PAFG305D = nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbDeleteOk          " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub

%>
