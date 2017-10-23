<%
Option Explicit
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : OPEN Ar 저장 업무 처리 ASP
'*  3. Program ID           : a3101mb2
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/11/12
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************




'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 
Err.Clear 

Call LoadBasisGlobalInf()

Dim lgIntFlgMode
Dim LngMaxRow
Dim LngMaxRow3
Dim iCommandSent
Dim AutoNum

Dim iPARG005																	'조회용 ComPlus Dll 사용 변수 
Dim iArrData
Dim iArrDept 
Dim iArrSpread
Dim iArrSpread3															
Dim iDealBpCd
Dim iPayBpCd
Dim iReportBpCd
Dim iAcctCd
Dim iChangeOrgId

const OpenArNo			=  0
Const OpenArDt			=  1
Const OpenInvDocNo		=  2
Const OpenInvDt			=  3
Const OpenBlDocNo		=  4
Const OpenBlDt			=  5
Const OpenRefNo         =  6
Const OpenDocCur		=  7
Const OpenXchRate		=  8
Const OpenArDueDt		=  9
Const OpenNetAmt        = 10
Const OpenNetLocAmt     = 11
Const OpenArAmt			= 12
Const OpenArLocAmt		= 13
Const OpenArType		= 14
Const OpenArSts			= 15
Const OpenConfFg		= 16
Const OpenRcptType		= 17
Const OpenRcptTerms		= 18
Const OpenDesc			= 19
Const OpenInsertUserId  = 20
Const OpenUpdtUserId    = 21
Const OpenCashAmt		= 22
Const OpenCashLocAmt	= 23
Const OpenPrRcptAmt		= 24
Const OpenPrRcptLocAmt	= 25
Const OpenPrRcptNo		= 26
Const OpenPayMethCd		= 27
Const OpenPayDur		= 28
Const OpenGldt			= 29
Const OpenProject		= 30

' -- 권한관리추가 
Const A114_I11_a_data_auth_data_BizAreaCd = 0
Const A114_I11_a_data_auth_data_internal_cd = 1
Const A114_I11_a_data_auth_data_sub_internal_cd = 2
Const A114_I11_a_data_auth_data_auth_usr_id = 3

Dim I11_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

Redim I11_a_data_auth(3)
I11_a_data_auth(A114_I11_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
I11_a_data_auth(A114_I11_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
I11_a_data_auth(A114_I11_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
I11_a_data_auth(A114_I11_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

'#########################################################################################################
'												 업무 처리 
'##########################################################################################################

	iChangeOrgId = UCase(Request("hOrgChangeId"))

	LngMaxRow    = CInt(Request("txtMaxRows"))									'☜: 최대 업데이트된 갯수 
	LngMaxRow3   = CInt(Request("txtMaxRows3"))									'☜: 최대 업데이트된 갯수 
	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 

	'-----------------------
	'Data manipulate area
	'-----------------------												    'Single 데이타 저장 

	Redim iarrdata(OpenProject)    

	iArrData(OpenArNo)           = Trim(Request("txtArNo"))
	iArrData(OpenArDt)		     = UNIConvDate(Request("txtArDt"))
	iArrData(OpenInvDt)		     = UNIConvDate(Request("txtInvDt"))
	iArrData(OpenInvDocNo)		 = Trim(Request("txtInvNo"))
	iArrData(OpenBlDocNo)		 = Trim(Request("txTblNo"))
	iArrData(OpenBlDt)			 = UNIConvDate(Request("txTblDt"))
	iArrData(OpenPayDur)		 = UNIConvNum(Request("txtPayDur"),0)
	iArrData(OpenPayMethCd)		 = Request("txtPayMethCd")
	iArrData(OpenArDueDt)	     = UNIConvDate(Request("txtDueDt"))
	iArrData(OpenRcptType)		 = Request("txtPayTypeCd")
	iArrData(OpenDocCur)		 = UCase(Trim(Request("txtDocCur")))
	iArrData(OpenRcptTerms)		 = Request("txtPaymTerms")	
	iArrData(OpenXchRate)		 = UNIConvNum(Request("txtXchRate"),0)
	iArrData(OpenArType)         = "NR"
	iArrData(OpenCashAmt)		 = UNIConvNum(Request("txtCashAmt"),0)
	iArrData(OpenCashLocAmt)	 = UNIConvNum(Request("txtCashLocAmt"),0)
	iArrData(OpenPrRcptAmt)		 = UNIConvNum(Request("txtPrRcptAmt"),0)
	iArrData(OpenPrRcptLocAmt)	 = UNIConvNum(Request("txtPrRcptLocAmt"),0)
	iArrData(OpenPrRcptNo)		 = Trim(Request("txtPrPaymNo"))
	iArrData(OpenDesc)		     = Request("txtDesc")
	iArrData(OpenProject)		 = Request("txtProject")
	iArrData(OpenArSts)          = "O"
	iArrData(OpenConfFg)		 = "U"
	

	iDealBpCd					 = UCase(Trim(Request("txtDealBpCd")))
	iPayBpCd					 = UCase(Trim(Request("txtPayBpCd")))
	iReportBpCd					 = UCase(Trim(Request("txtReportBpCd")))
	iAcctCd						 = Trim(Request("txtAcctcd"))
	
	Redim iArrDept(1)

	iArrDept(0) = iChangeOrgId
	iArrDept(1) = Trim(Request("txtDeptCd"))

	iArrSpread  = Request("txtSpread")
	iArrSpread3 = Request("txtSpread3")
	
	
	Set iPARG005 = Server.CreateObject("PARG005.cAMngOpenArSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	                       
	If lgIntFlgMode = OPMD_CMODE Then
		AutoNum = iPARG005.A_MANAGE_OPEN_AR_SVR(gStrGlobalCollection, "CREATE", ,gCurrency, , iArrDept, , _
											   iDealBpCd, iPayBpCd, iReportBpCd, _
											   iAcctcd, iArrData, iArrSpread, iArrSpread3, I11_a_data_auth)
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		AutoNum =  iPARG005.A_Manage_Open_Ar_Svr(gStrGlobalCollection, "UPDATE", ,gCurrency, ,iArrDept, , _
											   iDealBpCd, iPayBpCd, iReportBpCd, _
											   iAcctcd, iArrData, iArrSpread, iArrSpread3, I11_a_data_auth)
	End If

	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG005 = Nothing		
		Response.End 
	End If
	    
	Set iPARG005 = nothing

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.dbSaveOk(""" & AutoNum & """)" & vbcr
	Response.Write "</Script>" & vbcr
%>
