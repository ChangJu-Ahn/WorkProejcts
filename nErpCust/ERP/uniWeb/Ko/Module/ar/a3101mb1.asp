<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Open AR ��ȸ�ϴ� p/g
'*  3. Program ID           : a3101mb1	
'*  4. Program Name         :	
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : YOU SO EUN
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2.- ���Ǻ� 
'##########################################################################################################
																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
On Error Resume Next														'��: 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Const NonSalesAr = "NR"

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

'#########################################################################################################
'												2.1 ���� üũ 
'##########################################################################################################
If strMode = "" Then'
	Response.End
	Call HideStatusWnd		 	 	 
End If

'#########################################################################################################
'												2. ���� ó�� ����� 
'##########################################################################################################

'#########################################################################################################
'												2.1. ����, ��� ���� 
'##########################################################################################################
Dim iPARG005																	'�� : ��ȸ�� ComProxy Dll ��� ���� 

Dim ImportData 
Dim IntRows
Dim intCount
Dim IntCurSeq
Dim StrNextKey
Dim StrNextKeyDt1

Dim itxtArNo
Dim iarrData1
Dim iarrData2
Dim iStrData
Dim iLngRow
Dim iIntLoopCount
Dim LngMaxRow
Dim lgCurrency

Const OpenDealBpCd = 1
Const OpenDealBpNm = 2
Const OpenInvDocNo = 3
Const OpenPayBpCd = 4
Const OpenPayBpNm = 5
Const OpenInvDt = 6
Const OpenReportBpCd = 7 
Const OpenReportBpNm = 8
Const OpenBlDocNo = 9
Const OpenDeptCd = 10
Const OpenDeptNm = 11
Const OpenBlDt = 12
Const OpenAcctCd = 13
Const OpenAcctNm = 14
Const OpenPayDur = 15
Const OpenArDt = 16
Const OpenPayMethCd = 17
Const OpenPayMethNm = 18
Const OpenArDueDt = 19
Const OpenRcptType = 20
Const OpenRcptTypeNm =21
Const OpenDocCur = 22
Const OpenRcptTerms = 23
Const OpenXchRate = 24
Const OpenArType = 25
Const OpenVatAmt = 26
Const OpenVatLocAmt =27
Const OpenNetAmt = 28
Const OpenNetLocAmt = 29
Const OpenCashAmt = 30
Const OpenCashLocAmt = 31
Const OpenPrRcptAmt = 32
Const OpenPrRcptLocAmt= 33
Const OpenPrRcptNo = 34
Const OpenGlNo= 35
Const OpenArTotAmt = 36
Const OpenArTotLocAmt =37
Const OpenArAmt = 38
Const OpenArLocAmt = 39
Const OpenBalAmt = 40
Const OpenBalLocAmt = 41
Const OpenDesc = 42
Const OpenTempGlNo = 43
Const OpenProject = 45

Const ITEM_SEQ = 0
Const ACCT_CD  = 1
Const ACCT_NM  = 2
Const DEPT_CD  = 3
Const DEPT_NM  = 4
Const VAT_TYPE = 5
Const VAT_TYPE_NM = 6
Const TOT_AMT = 7
Const TOT_LOC_AMT = 8
Const ITEM_DESC  = 9


Const gIsShowLocal = "Y"

' -- ���Ѱ����߰� 
Const A336_I3_a_data_auth_data_BizAreaCd = 0
Const A336_I3_a_data_auth_data_internal_cd = 1
Const A336_I3_a_data_auth_data_sub_internal_cd = 2
Const A336_I3_a_data_auth_data_auth_usr_id = 3

Dim I3_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

Redim I3_a_data_auth(3)
I3_a_data_auth(A336_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
I3_a_data_auth(A336_I3_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
I3_a_data_auth(A336_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
I3_a_data_auth(A336_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
		
'#########################################################################################################
'												2.2. ��û ���� ó�� 
'##########################################################################################################

LngMaxRow = Cint(Request("txtMaxRows"))

'#########################################################################################################
'												2.3. ���� ó�� 
'##########################################################################################################

Set iPARG005 = Server.CreateObject("PARG005.cALkUpOpenArSvr")

If CheckSYSTEMError(Err,True) = True Then
	Set iPARG005 = nothing
	Response.End 
End If

Redim ImportData(1)
ImportData(0) = Trim(Request("txtArNo"))
ImportData(1) = "NR"

Call iPARG005.A_LOOKUP_OPEN_AR_SVR (gStrGlobalCollection, ImportData,iarrData1,iarrData2,I3_a_data_auth)

If CheckSYSTEMError(Err,True) = True Then
	Set iPARG005 = nothing		
	Response.End 
End If
    
Set iPARG005 = nothing 

'#########################################################################################################
'												2.4. HTML ��� ������ 
'##########################################################################################################
lgCurrency = ConvSPChars(iarrData1(OpenDocCur))

Response.Write "<Script Language=vbscript>   " & vbcr
Response.Write " With parent.frm1            " & vbcr														'��: ȭ�� ó�� ASP �� ��Ī�� 
Response.Write " .txtDealBpCd.Value			=   """ & ConvSPChars(iarrData1(OpenDealBpCd)) & """				" & vbcr
Response.Write " .txtDealBpNm.Value			=   """ & ConvSPChars(iarrData1(OpenDealBpNm)) & """				" & vbcr
Response.Write " .txtInvNo.Value			=   """ & ConvSPChars(iarrData1(OpenInvDocNo)) & """				" & vbcr
Response.Write " .txtPayBpCd.Value			=   """ & ConvSPChars(iarrData1(OpenPayBpCd)) & """					" & vbcr
Response.Write " .txtPayBpNm.Value	  		=   """ & ConvSPChars(iarrData1(OpenPayBpNm)) & """					" & vbcr
Response.Write " .txtInvDt.text				=   """ & UNIDateClientFormat(iarrData1(OpenInvDt)) & """   		" & vbcr
Response.Write " .txtReportBpCd.Value		=   """ & ConvSPChars(iarrData1(OpenReportBpCd)) & """				" & vbcr
Response.Write " .txtReportBpNm.Value		=   """ & ConvSPChars(iarrData1(OpenReportBpNm)) & """				" & vbcr
Response.Write " .txTblNo.Value				=   """ & ConvSPChars(iarrData1(OpenBlDocNo)) & """					" & vbcr	
Response.Write " .txtDeptCd.Value			=   """ & ConvSPChars(iarrData1(OpenDeptCd)) & """					" & vbcr
Response.Write " .txtDeptNm.Value			=   """ & ConvSPChars(iarrData1(OpenDeptNm)) & """					" & vbcr
Response.Write " .txTblDt.text				=   """ & UNIDateClientFormat(iarrData1(OpenBlDt)) & """			" & vbcr
Response.Write " .txtAcctCd.Value			=   """ & ConvSPChars(iarrData1(OpenAcctCd)) & """					" & vbcr
Response.Write " .txtAcctNm.Value			=   """ & ConvSPChars(iarrData1(OpenAcctNm)) & """					" & vbcr
Response.Write " .txtPayDur.Value			=	""" & UNINumClientFormat(iarrData1(OpenPayDur), 0, 0) & """		" & vbcr		
Response.Write " .txtArDt.text				=   """ & UNIDateClientFormat(iarrData1(OpenArDt)) & """			" & vbcr
Response.Write " .txtPayMethCd.Value		=   """ & ConvSPChars(iarrData1(OpenPayMethCd)) & """				" & vbcr			
Response.Write " .txtPayMethNm.Value		=   """ & ConvSPChars(iarrData1(OpenPayMethNm)) & """				" & vbcr			
Response.Write " .txtDueDt.text				=   """ & UNIDateClientFormat(iarrData1(OpenArDueDt)) & """			" & vbcr	
Response.Write " .txtPayTypeCd.Value		=   """ & ConvSPChars(iarrData1(OpenRcptType)) & """				" & vbcr			
Response.Write " .txtPayTypeNm.Value		=   """ & ConvSPChars(iarrData1(OpenRcptTypeNm)) & """				" & vbcr			
Response.Write " .txtDocCur.Value			=   """ & ConvSPChars(iarrData1(OpenDocCur)) & """					" & vbcr
Response.Write " .txtPaymTerms.Value		=   """ & ConvSPChars(iarrData1(OpenRcptTerms)) & """				" & vbcr			
'Response.Write " .cboArType.Value			=	""" & ConvSPChars(iarrData1(OpenArType)) & """					" & vbcr     '"NonSalesAr"

Response.Write " .txtCashAmt.Text			=   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenCashAmt),	lgCurrency,	ggAmtOfMoneyNo,	"X", "X") & """			" & vbcr
Response.Write " .txtPrRcptAmt.Text			=   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenPrRcptAmt),	lgCurrency,	ggAmtOfMoneyNo,	"X", "X") & """			" & vbcr
Response.Write " .txtArTotAmt.Text			=   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenArTotAmt),	lgCurrency,	ggAmtOfMoneyNo,	"X", "X") & """			" & vbcr
Response.Write " .txtArAmt.Text				=   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenArAmt),		lgCurrency,	ggAmtOfMoneyNo,	"X", "X") & """			" & vbcr
Response.Write " .txtBalAmt.Text			=   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenBalAmt),	lgCurrency,	ggAmtOfMoneyNo,	"X", "X") & """			" & vbcr

If gIsShowLocal <> "N" Then
	Response.Write " .txtXchRate.Text       =   """ & UNINumClientFormat(iarrData1(OpenXchRate),	ggExchRate.DecPoint,	0) & """											" & vbcr
	Response.Write " .txtVatAmt.Text        =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenVatAmt),		lgCurrency,	ggAmtOfMoneyNo,	"X",				"X") & """	" & vbcr
	Response.Write " .txtVatLocAmt.Text     =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenVatLocAmt),		gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,	"X") & """	" & vbcr
	Response.Write " .txtNetAmt.Text        =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenNetAmt),		lgCurrency,	ggAmtOfMoneyNo,	"X",				"X") & """	" & vbcr
	Response.Write " .txtNetLocAmt.Text     =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenNetLocAmt),		gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,	"X") & """	" & vbcr
	Response.Write " .txtCashLocAmt.Text    =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenCashLocAmt),	gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,	"X") & """	" & vbcr
	Response.Write " .txtPrRcptLocAmt.Text  =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenPrRcptLocAmt),	gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,	"X") & """	" & vbcr
	Response.Write " .txtArTotLocAmt.Text   =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenArTotLocAmt),	gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,	"X") & """	" & vbcr
	Response.Write " .txtArLocAmt.Text      =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenArLocAmt),		gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,	"X") & """	" & vbcr
	Response.Write " .txtBalLocAmt.Text     =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenBalLocAmt),		gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,	"X") & """	" & vbcr
Else
	Response.Write " .txtXchRate.Value      =   """ & UNINumClientFormat(iarrData1(OpenXchRate),	ggExchRate.DecPoint,	 0) & """										" & vbcr
	Response.Write " .txtVatAmt.Value       =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenVatAmt),		lgCurrency,	ggAmtOfMoneyNo,	"X",			"X") & """	" & vbcr
	Response.Write " .txtVatLocAmt.Value    =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenVatLocAmt),		gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,"X") & """	" & vbcr
	Response.Write " .txtNetAmt.Value       =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenNetAmt),		lgCurrency,	ggAmtOfMoneyNo,	"X",			"X") & """	" & vbcr
	Response.Write " .txtNetLocAmt.Value    =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenNetLocAmt),		gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,"X") & """	" & vbcr
	Response.Write " .txtCashLocAmt.Value   =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenCashLocAmt),	gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,"X") & """	" & vbcr
	Response.Write " .txtPrRcptLocAmt.Value =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenPrRcptLocAmt),	gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,"X") & """	" & vbcr
	Response.Write " .txtArTotLocAmt.Value  =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenArTotLocAmt),	gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,"X") & """	" & vbcr
	Response.Write " .txtArLocAmt.Value     =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenArLocAmt),		gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,"X") & """	" & vbcr
	Response.Write " .txtBalLocAmt.Value    =   """ & UNIConvNumDBToCompanyByCurrency(iarrData1(OpenBalLocAmt),		gCurrency,	ggAmtOfMoneyNo,	gLocRndPolicyNo,"X") & """	" & vbcr
End If

Response.Write " .txtPrPaymNo.Value			=   """ & ConvSPChars(iarrData1(OpenPrRcptNo)) & """				    " & vbcr				
Response.Write " .txtGlNo.Value				=   """ & ConvSPChars(iarrData1(OpenGlNo))     & """					" & vbcr
Response.Write " .txtTempGlNo.Value			=   """ & ConvSPChars(iarrData1(OpenTempGlNo)) & """					" & vbcr
Response.Write " .txtDesc.Value				=   """ & ConvSPChars(iarrData1(OpenDesc))     & """					" & vbcr
Response.Write " .txtProject.Value			=   """ & ConvSPChars(iarrData1(openProject))     & """					" & vbcr

Response.Write " End With						" & vbcr		    
Response.write "</Script>						" & vbcr  

iStrData = ""
iIntLoopCount = 0	

For iLngRow = 0 To UBound(iarrData2, 1)
	iStrData = iStrData & Chr(11) & iarrData2(iLngRow, ITEM_SEQ)
	iStrData = iStrData & Chr(11) & ConvSPChars(iarrData2(iLngRow, ACCT_CD))
	iStrData = iStrData & Chr(11) & ""		
	iStrData = iStrData & Chr(11) & ConvSPChars(iarrData2(iLngRow, ACCT_NM))
	iStrData = iStrData & Chr(11) & ConvSPChars(iarrData2(iLngRow, DEPT_CD))
	iStrData = iStrData & Chr(11) & ""		
	iStrData = iStrData & Chr(11) & ConvSPChars(iarrData2(iLngRow, DEPT_NM))
	iStrData = iStrData & Chr(11) & ConvSPChars(iarrData2(iLngRow, VAT_TYPE))
	iStrData = iStrData & Chr(11) & ""
	iStrData = iStrData & Chr(11) & ConvSPChars(iarrData2(iLngRow, VAT_TYPE_NM))		
	iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iarrData2(iLngRow, TOT_AMT), lgCurrency, ggAmtOfMoneyNo, "X",             "X")
	iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iarrData2(iLngRow, TOT_LOC_AMT), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
	iStrData = iStrData & Chr(11) & ConvSPChars(iarrData2(iLngRow,ITEM_DESC))
	iStrData = iStrData & Chr(11) & LngMaxRow + iLngRow + 1 
	iStrData = iStrData & Chr(11) & Chr(12)
Next

	Response.Write " <Script Language=vbscript>									" & vbCr
	Response.Write " With parent												" & vbCr
	Response.Write "	.ggoSpread.Source = .frm1.vspdData				" & vbcr
	Response.Write "	.ggoSpread.SSShowData """ & istrData & """		" & vbcr
	Response.Write "	.frm1.htxtArNo.value = """ & Request("txtArNo") & """	" & vbcr
	Response.Write "	.DbQueryOk												" & vbcr
    Response.Write " End With													" & vbCr
    Response.Write " </Script>												" & vbCr
	

%>
