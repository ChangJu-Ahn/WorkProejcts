
<%
Option Explicit
'**********************************************************************************************
'*  1. Module Name          : OPEN AP ���� ���� ó�� ASP
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************

																'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
																'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%																					'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd																	'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

On Error Resume Next																'��: 
Err.Clear 

Call LoadBasisGlobalInf()
'#########################################################################################################
'												1. ����, ��� ���� 
'##########################################################################################################

Dim lgIntFlgMode
Dim LngMaxRow
Dim LngMaxRow3
Dim ImportData1 
Dim AutoNum

Dim iPAPG005																		'����� ComPlus Dll ��� ���� 
Dim iChangeOrgId 
Dim iArrSpread
Dim iArrSpread3

Const ApNo			= 0
Const RefNo			= 1
Const BizAreaCd		= 2
Const PayBpCd		= 3
Const DealBpCd		= 4
Const ReportBpCd	= 5
Const DeptCd	    = 6
Const ChangeOrgId   = 7
Const ApDt		    = 8
Const DueDt			= 9
Const InvDocNo		= 10
Const InvDt		    = 11
Const LcDocNo		= 12
Const LcDt		    = 13
Const AcctCd		= 14
Const DocCur	    = 15
Const XchRate		= 16
Const BlDocNo		= 17
Const BlDt		    = 18
Const ApSts         = 19
Const ApType        = 20
Const PaymType		= 21
Const PaymTerms		= 22
Const PayDur		= 23
Const PayMeth		= 24
Const CashAmt		= 25
Const CashLocAmt	= 26
Const PrpaymAmt		= 27
Const PrpaymLocAmt	= 28
Const PrpaymNo		= 29
Const ConfFg		= 30
Const ApDesc		= 31
Const NetAmt        = 32
Const NetLocAmt     = 33
Const ApAmt         = 34
Const ApLocAmt      = 35
Const Gldt          = 36

	' -- ���Ѱ����߰� 
	Const A386_I4_a_data_auth_data_BizAreaCd = 0
	Const A386_I4_a_data_auth_data_internal_cd = 1
	Const A386_I4_a_data_auth_data_sub_internal_cd = 2
	Const A386_I4_a_data_auth_data_auth_usr_id = 3

	Dim I4_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

	Redim I4_a_data_auth(3)
	I4_a_data_auth(A386_I4_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I4_a_data_auth(A386_I4_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I4_a_data_auth(A386_I4_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I4_a_data_auth(A386_I4_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
'#########################################################################################################
'												2. ���� ó�� 
'##########################################################################################################

	iChangeOrgId = Trim(Request("hOrgChangeId"))

	LngMaxRow = CInt(Request("txtMaxRows"))											'��: �ִ� ������Ʈ�� ���� 
	LngMaxRow3 = CInt(Request("txtMaxRows3"))										'��: �ִ� ������Ʈ�� ���� 
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'��: ����� Create/Update �Ǻ� 

	ReDim ImportData1(Gldt)
	ImportData1(ApNo)            = Trim(Request("txtApNo"))
	ImportData1(RefNo)			 = Trim(Request("txtApNo"))
	ImportData1(BizAreaCd)		 = Trim(Request("txtReportBizCd"))
	ImportData1(PayBpCd)		 = Trim(Request("txtPayBpCd"))
	ImportData1(DealBpCd)		 = Trim(Request("txtDealBpCd"))
	ImportData1(ReportBpCd)		 = Trim(Request("txtReportBpCd"))
	ImportData1(DeptCd)			 = Trim(Request("txtDeptCd"))
	ImportData1(ChangeOrgId)     = iChangeOrgId
	ImportData1(ApDt)		     = UNIConvDate(Request("txtApDt"))
	ImportData1(DueDt)			 = UNIConvDate(Request("txtDueDt"))
	ImportData1(InvDocNo)		 = Trim(Request("txtInvNo"))
	ImportData1(InvDt)		     = UNIConvDate(Request("txtInvDt"))
	ImportData1(LcDocNo)		 = Trim(Request("txTlcNo"))
	ImportData1(LcDt)		     = UNIConvDate(Request("txtLcDt"))
	ImportData1(AcctCd)		     = Trim(Request("txtAcctCd"))
	ImportData1(DocCur)		     = Trim(Request("txtDocCur"))
	ImportData1(XchRate)		 = UNIConvNum(Request("txtXchRate"),0)
	ImportData1(BlDocNo)		 = Trim(Request("txTblNo"))
	ImportData1(BlDt)		     = UNIConvDate(Request("txTblDt"))
	ImportData1(ApSts)           = "O"
	ImportData1(ApType)          = "NP"
	ImportData1(PaymType)		 = Request("txtPayTypeCd")
	ImportData1(PaymTerms)		 = Request("txtPaymTerms")
	ImportData1(PayDur)			 = UNIConvNum(Request("txtPayDur"),0)
	ImportData1(PayMeth)		 = Request("txtPayMethCd")
	ImportData1(CashAmt)		 = UNIConvNum(Request("txtCashAmt"),0)
	ImportData1(CashLocAmt)		 = UNIConvNum(Request("txtCashLocAmt"),0)	
	ImportData1(PrpaymAmt)		 = UNIConvNum(Request("txtPrPaymAmt"),0)
	ImportData1(PrpaymLocAmt)	 = UNIConvNum(Request("txtPrPaymLocAmt"),0)	
	ImportData1(PrpaymNo)		 = Trim(Request("txtPrPaymNo"))
	ImportData1(ConfFg)		     = "U"
	ImportData1(ApDesc)		     = Request("txtDesc") 
	ImportData1(GlDt)		     = UNIConvDate(Request("txTGlDt"))	

	Set iPAPG005 = Server.CreateObject("PAPG005.cAMngOpenApSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If



	iArrSpread = Request("txtSpread")
	iArrSpread3 = Request("txtSpread3")
	
	If lgIntFlgMode = OPMD_CMODE Then
		AutoNum = iPAPG005.A_MANAGE_OPEN_AP_SVR (gStrGlobalCollection, "CREATE", , ImportData1, _
												    gCurrency, iArrSpread, iArrSpread3, I4_a_data_auth)
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		AutoNum = iPAPG005.A_MANAGE_OPEN_AP_SVR (gStrGlobalCollection, "UPDATE", , ImportData1, _
													gCurrency, iArrSpread, iArrSpread3, I4_a_data_auth)
	End If
	    
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG005 = Nothing		
		Response.End 
	End If
	    
	Set iPAPG005 = Nothing 

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.dbSaveOk(""" & AutoNum & """)" & vbcr
	Response.Write "</Script>" & vbcr

%>

