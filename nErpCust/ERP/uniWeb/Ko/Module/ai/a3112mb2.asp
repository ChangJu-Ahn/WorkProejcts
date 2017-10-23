
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : ���� OPEN AR ���� ���� ó�� ASP
'*  3. Program ID           : A3112mb2
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2003/01/07
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************

Response.Expires = -1															'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True															'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%																				'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

	On Error Resume Next														'��: 
	Err.Clear 

	Call LoadBasisGlobalInf()

	Dim lgIntFlgMode
	Dim AutoNum

	Dim iPARG005																'��ȸ�� ComPlus Dll ��� ���� 
	Dim iArrData
	Dim iArrDept
	Dim iReportBizCd
	Dim iDealBpCd
	Dim iPayBpCd
	Dim iReportBpCd
	Dim iAcctCd
	Dim iChangeOrgId													
	Dim ImportTypeTransType

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
	Const OpenGlDt			= 29
	Const OpenProject		= 30	

	iChangeOrgId = UCase(Request("hOrgChangeId"))
	lgIntFlgMode = Cint(Request("txtFlgMode")) 										'��: ����� Create/Update �Ǻ� 
	
'#########################################################################################################
'												 ���� ó�� 
'##########################################################################################################

	'-----------------------
	'Data manipulate area
	'-----------------------												    'Single ����Ÿ ���� 
	ImportTypeTransType = "AR008"

	Redim iarrdata(OpenProject)    

	iArrData(OpenArDt)		     = UNIConvDate(Request("txtArDt"))
	iArrData(OpenInvDt)		     = UNIConvDate(Request("txtInvDt"))
	iArrData(OpenInvDocNo)		 = Trim(Request("txtInvNo"))
	iArrData(OpenBlDocNo)		 = Trim(Request("txTblNo"))
	iArrData(OpenBlDt)			 = UNIConvDate(Request("txTblDt"))
	iArrData(OpenPayDur)		 = UNIConvNum(Request("txtPayDur"),0)
	iArrData(OpenPayMethCd)		 = Request("txtPayMethCd")
	iArrData(OpenArDueDt)	     = UNIConvDate(Request("txtDueDt"))
	iArrData(OpenRcptType)		 = Request("txtPayTypeCd")
	iArrData(OpenDocCur)		 = Request("txtDocCur")
	iArrData(OpenRcptTerms)		 = Request("txtPaymTerms")
	iArrData(OpenXchRate)		 = UNIConvNum(Request("txtXchRate"),0)
	iArrData(OpenArType)         = "NT"
	iArrData(OpenCashAmt)		 = UNIConvNum(Request("txtCashAmt"),0)
	iArrData(OpenCashLocAmt)	 = UNIConvNum(Request("txtCashLocAmt"),0)	
	iArrData(OpenNetAmt)		 = UNIConvNum(Request("txtNetAmt"),0)
	iArrData(OpenNetLocAmt)		 = UNIConvNum(Request("txtNetLocAmt"),0)	
	iArrData(OpenPrRcptAmt)		 = 0
	iArrData(OpenPrRcptLocAmt)	 = 0	
	iArrData(OpenPrRcptNo)		 = ""
	iArrData(OpenDesc)		     = Request("txtDesc") 
	iArrData(OpenArSts)          = "O"
	iArrData(OpenConfFg)		 = "U"
	iArrData(OpenArNo)           = Trim(Request("txtArNo"))
	iArrData(OpenGlDt)           = UNIConvDate(Request("txtGlDt"))
	iArrData(OpenProject)		 = Request("txtProject")

	Redim iArrDept(1)

	iArrDept(0)					= iChangeOrgId
	iArrDept(1)					= Trim(Request("txtDeptCd"))

	iReportBizCd				= Trim(Request("txtReportBizCd"))
	iDealBpCd					= Trim(Request("txtDealBpCd"))
	iPayBpCd					= Trim(Request("txtPayBpCd"))
	iReportBpCd					= Trim(Request("txtReportBpCd"))
	iAcctCd						= Trim(Request("txtAcctcd"))

	Set iPARG005 = Server.CreateObject("PARG005.cAMngOpenArSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	   
	If lgIntFlgMode = OPMD_CMODE Then
		AutoNum = iPARG005.A_Manage_Open_Ar_Svr(gStrGlobalCollection, "CREATE", ImportTypeTransType, gCurrency, , _
												iArrDept, iReportBizCd, iDealBpCd, iPayBpCd, iReportBpCd, _
												iAcctcd, iArrData)
	ElseIf lgIntFlgMode = OPMD_UMODE Then	
	    AutoNum =  iPARG005.A_Manage_Open_Ar_Svr(gStrGlobalCollection, "UPDATE", ImportTypeTransType, gCurrency, , _
												iArrDept, iReportBizCd, iDealBpCd, iPayBpCd, iReportBpCd, _
												iAcctcd, iArrData)
	End If

	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG005 = Nothing		
		Response.End 
	End If
	    
	Set iPARG005 = Nothing

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.dbSaveOk(""" & AutoNum & """)" & vbcr
	Response.Write "</Script>" & vbcr
%>
