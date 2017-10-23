<%
Option Explicit
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : ���� Open Ap ��ȸ�ϴ� p/g
'*  3. Program ID           : a4110mb1	
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

Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.

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
'On Error Resume Next														'��: 
'Err.Clear 

Call HideStatusWnd()														'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

'#########################################################################################################
'												2.1 ���� üũ 
'##########################################################################################################
If strMode = "" Then'
	Response.End
	Call HideStatusWnd()		 		 
End If

'#########################################################################################################
'												2. ���� ó�� ����� 
'##########################################################################################################

'#########################################################################################################
'												2.1. ����, ��� ���� 
'##########################################################################################################
Dim iPAPG005																'�� : ��ȸ�� ComPlus Dll ��� ���� 
Dim ImportData 
Dim ExportData1
Dim ExportData2
Dim lgCurrency

Const acctcd		= 0
Const acctnm		= 1
Const BpCd			= 2
Const BpNm			= 3
Const apno			= 4
Const apdt			= 5
Const invdocno		= 6
Const invdt			= 7
Const BlDocNo		= 8
Const bldt			= 9
Const refno			= 10
Const doccur		= 11
Const xchrate		= 12
Const DueDt			= 13
Const NetAmt		= 14
Const Netlocamt		= 15
Const VatAmt		= 16
Const VatLocAmt		= 17
Const ApAmt			= 18
Const ApLocAmt		= 19
Const ApType		= 20
Const PaymType		= 21
Const Paymterms		= 22
Const ApDesc		= 23
Const LcDocNo		= 24
Const ApLcDt		= 25
Const CashAmt		= 26
Const CashLocAmt	= 27
Const PrPaymamt		= 28
Const PrPaymlocamt	= 29
Const PrPaymno		= 30
Const BalAmt		= 31
Const BalLocAmt		= 32
Const TotApAmt		= 33
Const TotApLocAmt	= 34
Const PayMeth		= 35
Const PayDur		= 36
Const tempglno		= 37
Const DeptCd		= 38
Const DeptNm		= 39
Const ReportBpCd	= 40
Const ReportBpNm	= 41
Const PayBpCd		= 42
Const PayBpNm		= 43
Const bizareacd		= 44
Const bizareanm		= 45
Const PaymTypeNm	= 46
Const PayMethNm		= 47
Const ConfFg		= 48
Const GlNo			= 49
Const Gldt			= 50

Const gIsShowLocal = "Y"

'#########################################################################################################
'												2.2. ��û ���� ó�� 
'##########################################################################################################

'#########################################################################################################
'												2.3. ���� ó�� 
'##########################################################################################################
	Redim ImportData(1)
	ImportData(0) = Trim(Request("txtApNo"))
	importData(1) = "NT"

	Set iPAPG005 = Server.CreateObject("PAPG005.cALkUpOpenApSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	    
	call iPAPG005.A_LOOKUP_OPEN_AP_SVR(gStrGlobalCollection, ImportData ,ExportData1,ExportData2)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG005 = Nothing		
		Response.End 
	End If
	    
	Set iPAPG005 = Nothing 

	lgCurrency = ConvSPChars(ExportData1(doccur))

'#########################################################################################################
'												2.4. HTML ��� ������ 
'##########################################################################################################
	Response.Write "<Script Language=vbscript>   " & vbcr
	Response.Write " With parent.frm1            " & vbcr														'��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write "	.txtDealBpCd.Value      =   """ & ConvSPChars(ExportData1(BpCd)) & """			" & vbcr
	Response.Write "	.txtDealBpNm.Value      =   """ & ConvSPChars(ExportData1(BpNm)) & """			" & vbcr
	Response.Write "	.txtDeptCd.Value        =   """ & ConvSPChars(ExportData1(DeptCd)) & """		" & vbcr
	Response.Write "	.txtDeptNm.Value        =   """ & ConvSPChars(ExportData1(DeptNm)) & """		" & vbcr
	Response.Write "	.txtPayBpCd.Value       =   """ & ConvSPChars(ExportData1(PayBpCd)) & """		" & vbcr
	Response.Write "	.txtPayBpNm.Value		=   """ & ConvSPChars(ExportData1(PayBpNm)) & """		" & vbcr
	Response.Write "	.txtReportBpCd.Value    =   """ & ConvSPChars(ExportData1(ReportBpCd)) & """	" & vbcr
	Response.Write "	.txtReportBpNm.Value	=   """ & ConvSPChars(ExportData1(ReportBpNm)) & """	" & vbcr	
	Response.Write "	.txtApDt.text           =   """ & UNIDateClientFormat(ExportData1(ApDt)) & """	" & vbcr
	Response.Write "	.txtInvNo.Value         =   """ & ConvSPChars(ExportData1(InvDocNo)) & """		" & vbcr
	Response.Write "	.txtDueDt.text          =   """ & UNIDateClientFormat(ExportData1(DueDt)) & """ " & vbcr	
	Response.Write "	.txtInvDt.text          =   """ & UNIDateClientFormat(ExportData1(InvDt)) & """ " & vbcr
	Response.Write "	.txtAcctCd.Value        =   """ & ConvSPChars(ExportData1(AcctCd)) & """		" & vbcr
	Response.Write "	.txtAcctNm.Value        =   """ & ConvSPChars(ExportData1(AcctNm)) & """		" & vbcr
	Response.Write "	.txTblNo.Value          =   """ & ConvSPChars(ExportData1(BlDocNo)) & """		" & vbcr	
	Response.Write "	.txTblDt.text           =   """ & UNIDateClientFormat(ExportData1(BlDt)) & """	" & vbcr
	Response.Write "	.txtPaymTerms.Value     =   """ & ConvSPChars(ExportData1(PaymTerms)) & """		" & vbcr
	Response.Write "	.txTlcNo.Value          =   """ & ConvSPChars(ExportData1(LcDocNo)) & """		" & vbcr
'	Response.Write "	.cboApType.Value        =   """ & ConvSPChars(ExportData1(ApType)) & """		" & vbcr	
	Response.Write "	.txtLcDt.text           =   """ & UNIDateClientFormat(ExportData1(ApLcDt)) & """" & vbcr
	Response.Write "	.txtPayDur.Value		=	""" & UniConvNum(ExportData1(PayDur), 0) & """		" & vbcr		
	Response.Write "	.txtPayMethCd.Value		=   """ & ConvSPChars(ExportData1(PayMeth)) & """		" & vbcr			
	Response.Write "	.txtPayMethNm.Value		=   """ & ConvSPChars(ExportData1(PayMethNm)) & """		" & vbcr			
	Response.Write "	.txtPayTypeCd.Value     =   """ & ConvSPChars(ExportData1(PaymType)) & """		" & vbcr	
	Response.Write "	.txtPayTypeNm.Value     =   """ & ConvSPChars(ExportData1(PaymTypeNm)) & """	" & vbcr		
	Response.Write "	.txtDocCur.Value        =   """ & ConvSPChars(ExportData1(DocCur)) & """		" & vbcr

	Response.Write "    .txtNetAmt.Text         =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(NetAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbcr
	Response.Write "    .txtBalAmt.Text         =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(BalAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbcr

	If gIsShowLocal <> "N" Then		
		Response.Write "	.txtXchRate.Text    =   """ & UNINumClientFormat(ExportData1(XchRate), ggExchRate.DecPoint, 0) & """		" & vbcr
		Response.Write "	.txtNetLocAmt.Text  =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(NetLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """	" & vbcr
		Response.Write "	.txtBalLocAmt.Text  =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(BalLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") & """" & vbcr
	Else
		Response.Write "	.txtXchRate.Value   =   """ & UNINumClientFormat(ExportData1(XchRate), ggExchRate.DecPoint, 0) & """		" & vbcr
		Response.Write "	.txtNetLocAmt.Value =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(NetLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """	" & vbcr
		Response.Write "	.txtBalLocAmt.Value =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(BalLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") & """" & vbcr
	End IF
		
	Response.Write "	.txtGlNo.Value          =   """ & ConvSPChars(ExportData1(GlNo)) & """			" & vbcr
	Response.Write "	.txtGldt.Text           =   """ & UNIDateClientFormat(ExportData1(Gldt)) & """	" & vbcr									
	Response.Write "	.txtDesc.Value	        =   """ & ConvSPChars(ExportData1(ApDesc)) & """		" & vbcr				
	Response.Write "	.htxtApNo.value			=   """ & Request("txtApNo") & """						" & vbcr

	Response.Write " End With					 " & vbcr		    
	Response.Write " Parent.DbQueryOk			 " & vbcr
	Response.write "</Script>				     " & vbcr  
%>    


